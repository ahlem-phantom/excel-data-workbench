package app.model;

import java.util.List;
import org.apache.poi.ss.util.CellRangeAddress;

/**
 * Immutable model holding a fully-parsed Excel sheet.
 *
 * cells[row][col]  – all cell values as strings (merged regions are propagated
 *                    so every cell in a merged block holds the same value).
 * headerRows       – number of top rows that define the column structure.
 *                    Detected automatically: 1 for flat headers, 2+ for grouped/multi-level headers.
 */
public class ExcelData {

    public final String[][] cells;
    public final int headerRows;
    public final int columnCount;
    public final int rowCount;
    public final List<CellRangeAddress> mergedRegions;

    public ExcelData(String[][] cells, int headerRows,
                     int columnCount, int rowCount,
                     List<CellRangeAddress> mergedRegions) {
        this.cells         = cells;
        this.headerRows    = headerRows;
        this.columnCount   = columnCount;
        this.rowCount      = rowCount;
        this.mergedRegions = mergedRegions;
    }

    /** Number of data rows (total rows minus header rows). */
    public int dataRowCount() {
        return Math.max(0, rowCount - headerRows);
    }
}
