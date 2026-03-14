package app.service;

import java.io.File;
import java.io.FileInputStream;
import java.util.List;

import app.model.ExcelData;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Parses any .xlsx file into an {@link ExcelData} model.
 *
 * Header detection algorithm
 * --------------------------
 * The importer inspects the merged regions declared in the spreadsheet
 * to automatically determine how many top rows form the column header:
 *
 *   1. If any merged region starts at row 0 and spans multiple rows
 *      (firstRow=0, lastRow>0), the header depth = lastRow + 1.
 *      Example: a 3-row grouped header → headerRows = 3.
 *
 *   2. If row 0 has horizontal merges (columns only, same row),
 *      headerRows = 2  (row 0 = group labels, row 1 = sub-columns).
 *
 *   3. Otherwise headerRows = 1  (simple single-row header).
 *
 * Merged cell propagation
 * -----------------------
 * Apache POI stores the value only in the top-left cell of a merged block.
 * After reading, this class copies that value to every cell in the block so
 * that downstream consumers always get the correct text regardless of which
 * cell they query.
 */
public class ExcelImporter {

    /**
     * Loads the first sheet of the given .xlsx file.
     *
     * @param path absolute path to the Excel file
     * @return fully parsed {@link ExcelData}
     */
    public ExcelData load(String path) throws Exception {
        try (FileInputStream fis = new FileInputStream(new File(path));
             XSSFWorkbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0);
            List<CellRangeAddress> mergedRegions = sheet.getMergedRegions();
            DataFormatter formatter = new DataFormatter();

            // Determine grid dimensions
            int rowCount = sheet.getLastRowNum() + 1;
            int colCount = 0;
            for (Row row : sheet) {
                colCount = Math.max(colCount, row.getLastCellNum());
            }
            if (rowCount == 0 || colCount <= 0) {
                return new ExcelData(new String[0][0], 1, 0, 0, mergedRegions);
            }

            // Read every cell into a 2-D String array
            String[][] cells = new String[rowCount][colCount];
            for (Row row : sheet) {
                for (Cell cell : row) {
                    int c = cell.getColumnIndex();
                    if (c < colCount) {
                        cells[row.getRowNum()][c] = formatter.formatCellValue(cell).trim();
                    }
                }
            }

            // Propagate merged-region values to every cell in the block
            fillMergedRegions(cells, mergedRegions, rowCount, colCount);

            int headerRows = detectHeaderRows(mergedRegions);

            return new ExcelData(cells, headerRows, colCount, rowCount, mergedRegions);
        }
    }

    // -----------------------------------------------------------------------
    // Private helpers
    // -----------------------------------------------------------------------

    /**
     * Detects how many header rows the sheet has by inspecting its merged regions.
     *
     * Rule 1 – Multi-row span starting at row 0 → headerRows = max(lastRow)+1.
     * Rule 2 – Horizontal-only merge at row 0     → headerRows = 2.
     * Default – No qualifying merges               → headerRows = 1.
     */
    private int detectHeaderRows(List<CellRangeAddress> regions) {
        int maxCoveredRow = 0;
        boolean hasHorizontalTopMerge = false;

        for (CellRangeAddress r : regions) {
            if (r.getFirstRow() == 0) {
                if (r.getLastRow() > 0) {
                    // Spans multiple rows (vertical + horizontal)
                    maxCoveredRow = Math.max(maxCoveredRow, r.getLastRow());
                } else if (r.getFirstColumn() != r.getLastColumn()) {
                    // Horizontal-only merge in row 0
                    hasHorizontalTopMerge = true;
                }
            }
        }

        if (maxCoveredRow > 0)      return maxCoveredRow + 1;
        if (hasHorizontalTopMerge)  return 2;
        return 1;
    }

    /**
     * Copies the value stored in the top-left cell of each merged block to
     * every other cell within that block.
     */
    private void fillMergedRegions(String[][] cells,
                                   List<CellRangeAddress> regions,
                                   int rowCount, int colCount) {
        for (CellRangeAddress r : regions) {
            String origin = safeGet(cells, r.getFirstRow(), r.getFirstColumn());
            if (origin == null) origin = "";
            for (int row = r.getFirstRow(); row <= r.getLastRow() && row < rowCount; row++) {
                for (int col = r.getFirstColumn(); col <= r.getLastColumn() && col < colCount; col++) {
                    if (cells[row][col] == null || cells[row][col].isEmpty()) {
                        cells[row][col] = origin;
                    }
                }
            }
        }
    }

    private String safeGet(String[][] cells, int row, int col) {
        if (row < cells.length && col < cells[row].length) return cells[row][col];
        return null;
    }
}
