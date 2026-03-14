package app.service;

import java.io.File;
import java.io.FileOutputStream;

import app.model.ExcelData;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;

/**
 * Writes an {@link ExcelData} model back to an .xlsx file.
 *
 * The output preserves:
 *   - All merged regions from the original import
 *   - Styled header rows (dark-blue background, white bold Calibri)
 *   - Auto-sized column widths
 */
public class ExcelExporter {

    public void export(ExcelData data, String outputPath) throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {

            XSSFSheet sheet = workbook.createSheet("Sheet1");

            CellStyle headerStyle = buildHeaderStyle(workbook);
            CellStyle dataStyle   = buildDataStyle(workbook);

            // Write all rows (headers + data)
            for (int r = 0; r < data.rowCount; r++) {
                Row row = sheet.createRow(r);
                row.setHeightInPoints(r < data.headerRows ? 22f : 16f);

                CellStyle style = r < data.headerRows ? headerStyle : dataStyle;

                for (int c = 0; c < data.columnCount; c++) {
                    Cell cell = row.createCell(c);
                    String val = data.cells[r][c];
                    cell.setCellValue(val != null ? val : "");
                    cell.setCellStyle(style);
                }
            }

            // Restore merged regions (only the top-left cell holds the value in XLSX)
            for (CellRangeAddress region : data.mergedRegions) {
                if (region.getLastRow() < data.rowCount && region.getLastColumn() < data.columnCount) {
                    sheet.addMergedRegion(region);
                }
            }

            // Auto-fit column widths (capped to avoid absurdly wide columns)
            for (int c = 0; c < data.columnCount; c++) {
                sheet.autoSizeColumn(c);
                int width = sheet.getColumnWidth(c);
                if (width > 20_000) sheet.setColumnWidth(c, 20_000);
            }

            String filePath = outputPath.endsWith(".xlsx") ? outputPath : outputPath + ".xlsx";
            try (FileOutputStream fos = new FileOutputStream(new File(filePath))) {
                workbook.write(fos);
            }
        }
    }

    // -----------------------------------------------------------------------
    // Style builders
    // -----------------------------------------------------------------------

    private CellStyle buildHeaderStyle(XSSFWorkbook wb) {
        XSSFCellStyle style = wb.createCellStyle();
        style.setFillForegroundColor(new XSSFColor(new byte[]{0x1e, 0x50, (byte) 0xa0}, null));
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);

        XSSFFont font = wb.createFont();
        font.setColor(new XSSFColor(new byte[]{(byte) 0xff, (byte) 0xff, (byte) 0xff}, null));
        font.setBold(true);
        font.setFontName("Calibri");
        font.setFontHeightInPoints((short) 11);
        style.setFont(font);
        return style;
    }

    private CellStyle buildDataStyle(XSSFWorkbook wb) {
        CellStyle style = wb.createCellStyle();
        style.setAlignment(HorizontalAlignment.LEFT);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        return style;
    }
}
