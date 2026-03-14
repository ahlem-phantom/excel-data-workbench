package app.ui;

import java.awt.BorderLayout;
import java.awt.Color;
import java.awt.Dimension;
import java.awt.FlowLayout;
import java.awt.Font;
import java.awt.Rectangle;
import java.awt.event.ActionEvent;
import java.awt.event.KeyEvent;
import java.util.ArrayDeque;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Comparator;
import java.util.Deque;
import java.util.HashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Objects;

import javax.swing.BorderFactory;
import javax.swing.Box;
import javax.swing.BoxLayout;
import javax.swing.JButton;
import javax.swing.JCheckBoxMenuItem;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JMenu;
import javax.swing.JMenuBar;
import javax.swing.JMenuItem;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JScrollPane;
import javax.swing.JTable;
import javax.swing.JTextField;
import javax.swing.KeyStroke;
import javax.swing.SwingConstants;
import javax.swing.event.TableModelEvent;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.table.DefaultTableCellRenderer;
import javax.swing.table.DefaultTableModel;
import javax.swing.table.JTableHeader;
import javax.swing.table.TableRowSorter;

import org.apache.poi.ss.util.CellRangeAddress;

import app.model.ExcelData;
import app.service.ExcelExporter;
import app.service.ExcelImporter;
import app.ui.components.GradientPanel;
import app.ui.components.StyledButton;

/**
 * Interactive workspace for editing, searching, navigating, and exporting Excel data.
 */
public class MainFrame extends JFrame {

    private static final long serialVersionUID = 1L;

    private static final Color NAVY = new Color(17, 59, 96);
    private static final Color TEAL = new Color(20, 146, 138);
    private static final Color SKY = new Color(237, 245, 252);
    private static final Color INK = new Color(35, 44, 58);

    private final ExcelImporter importer = new ExcelImporter();
    private final ExcelExporter exporter = new ExcelExporter();

    private ExcelData currentData;
    private String currentFilePath;

    private EditableTableModel tableModel;
    private JTable dataTable;
    private TableRowSorter<EditableTableModel> sorter;

    private JTextField searchField;
    private JLabel statusLabel;
    private JLabel fileLabel;
    private JPanel tableHostPanel;

    private boolean suppressSync;
    private int lastFoundModelRow = -1;
    private int lastFoundModelCol = -1;
    private String lastQuery = "";

    private final Deque<DeletedRowsSnapshot> undoStack = new ArrayDeque<DeletedRowsSnapshot>();

    private boolean editLocked;

    public MainFrame(String username) {
        setTitle("Excel Insight Studio - Workspace - " + username);
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setExtendedState(JFrame.MAXIMIZED_BOTH);
        setMinimumSize(new Dimension(1120, 700));

        setJMenuBar(buildMenuBar());
        setContentPane(buildRootPanel(username));
    }

    private JPanel buildRootPanel(String username) {
        JPanel root = new JPanel(new BorderLayout());
        root.setBackground(SKY);

        root.add(buildHeaderPanel(username), BorderLayout.NORTH);

        tableHostPanel = new JPanel(new BorderLayout());
        tableHostPanel.setBackground(SKY);
        tableHostPanel.setBorder(BorderFactory.createEmptyBorder(16, 16, 10, 16));
        tableHostPanel.add(buildEmptyStatePanel(), BorderLayout.CENTER);
        root.add(tableHostPanel, BorderLayout.CENTER);

        statusLabel = new JLabel("Ready. Import an Excel file to begin.");
        statusLabel.setFont(new Font("Segoe UI", Font.PLAIN, 13));
        statusLabel.setForeground(new Color(78, 90, 110));
        JPanel statusWrap = new JPanel(new BorderLayout());
        statusWrap.setBackground(Color.WHITE);
        statusWrap.setBorder(BorderFactory.createEmptyBorder(8, 14, 10, 14));
        statusWrap.add(statusLabel, BorderLayout.WEST);
        root.add(statusWrap, BorderLayout.SOUTH);

        return root;
    }

    private JPanel buildHeaderPanel(String username) {
        GradientPanel top = new GradientPanel(NAVY, TEAL);
        top.setLayout(new BorderLayout(14, 10));
        top.setBorder(BorderFactory.createEmptyBorder(16, 18, 16, 18));

        JPanel left = new JPanel();
        left.setOpaque(false);
        left.setLayout(new BoxLayout(left, BoxLayout.Y_AXIS));

        JLabel title = new JLabel("Data Editing Studio");
        title.setForeground(Color.WHITE);
        title.setFont(new Font("Segoe UI", Font.BOLD, 30));

        JLabel subtitle = new JLabel("Edit, search, jump to results, delete, undo, and export.");
        subtitle.setForeground(new Color(215, 239, 248));
        subtitle.setFont(new Font("Segoe UI", Font.PLAIN, 15));

        fileLabel = new JLabel("User: " + username + " | No file loaded");
        fileLabel.setForeground(new Color(224, 244, 240));
        fileLabel.setFont(new Font("Segoe UI", Font.BOLD, 12));

        left.add(title);
        left.add(Box.createVerticalStrut(4));
        left.add(subtitle);
        left.add(Box.createVerticalStrut(8));
        left.add(fileLabel);

        JPanel right = new JPanel();
        right.setOpaque(false);
        right.setLayout(new BoxLayout(right, BoxLayout.Y_AXIS));

        JPanel searchRow = new JPanel(new FlowLayout(FlowLayout.RIGHT, 8, 0));
        searchRow.setOpaque(false);
        searchField = new JTextField(24);
        searchField.setFont(new Font("Segoe UI", Font.PLAIN, 14));
        searchField.setBorder(BorderFactory.createCompoundBorder(
                BorderFactory.createLineBorder(new Color(210, 227, 238)),
                BorderFactory.createEmptyBorder(6, 10, 6, 10)));
        searchField.addActionListener(this::onFindNext);

        JButton findButton = createActionButton("Find Next", new Color(255, 143, 0), Color.WHITE);
        findButton.addActionListener(this::onFindNext);

        JButton gotoButton = createActionButton("Go To Row", new Color(0, 150, 136), Color.WHITE);
        gotoButton.addActionListener(this::onGotoRow);

        searchRow.add(searchField);
        searchRow.add(findButton);
        searchRow.add(gotoButton);

        JPanel actionRow = new JPanel(new FlowLayout(FlowLayout.RIGHT, 8, 8));
        actionRow.setOpaque(false);

        JButton importButton = createActionButton("Import", new Color(30, 136, 229), Color.WHITE);
        importButton.addActionListener(this::onImport);

        JButton exportButton = createActionButton("Export", new Color(56, 142, 60), Color.WHITE);
        exportButton.addActionListener(this::onExport);

        JButton deleteButton = createActionButton("Delete Row", new Color(229, 57, 53), Color.WHITE);
        deleteButton.addActionListener(this::onDeleteSelectedRows);

        JButton undoButton = createActionButton("Undo Delete", new Color(244, 81, 30), Color.WHITE);
        undoButton.addActionListener(this::onUndoDelete);

        JButton editLockButton = createActionButton("Lock/Unlock Edit", new Color(123, 31, 162), Color.WHITE);
        editLockButton.addActionListener(this::onToggleEditLock);

        actionRow.add(importButton);
        actionRow.add(exportButton);
        actionRow.add(deleteButton);
        actionRow.add(undoButton);
        actionRow.add(editLockButton);

        right.add(searchRow);
        right.add(actionRow);

        top.add(left, BorderLayout.WEST);
        top.add(right, BorderLayout.EAST);
        return top;
    }

    private JPanel buildEmptyStatePanel() {
        JPanel empty = new JPanel(new BorderLayout());
        empty.setBackground(Color.WHITE);
        empty.setBorder(BorderFactory.createCompoundBorder(
                BorderFactory.createLineBorder(new Color(224, 232, 241)),
                BorderFactory.createEmptyBorder(24, 24, 24, 24)));

        JLabel heading = new JLabel("No Data Loaded", SwingConstants.CENTER);
        heading.setFont(new Font("Segoe UI", Font.BOLD, 26));
        heading.setForeground(new Color(49, 64, 88));

        JLabel note = new JLabel("Import a .xlsx file to start editing and searching.", SwingConstants.CENTER);
        note.setFont(new Font("Segoe UI", Font.PLAIN, 15));
        note.setForeground(new Color(103, 120, 145));

        empty.add(heading, BorderLayout.CENTER);
        empty.add(note, BorderLayout.SOUTH);
        return empty;
    }

    private JButton createActionButton(String text, Color bg, Color fg) {
        StyledButton button = new StyledButton(text, bg, fg);
        button.setFont(new Font("Segoe UI", Font.BOLD, 13));
        return button;
    }

    private JMenuBar buildMenuBar() {
        JMenuBar bar = new JMenuBar();

        JMenu file = new JMenu("File");

        JMenuItem importItem = new JMenuItem("Import .xlsx");
        importItem.setAccelerator(KeyStroke.getKeyStroke(KeyEvent.VK_O, KeyEvent.CTRL_DOWN_MASK));
        importItem.addActionListener(this::onImport);

        JMenuItem exportItem = new JMenuItem("Export Current View");
        exportItem.setAccelerator(KeyStroke.getKeyStroke(KeyEvent.VK_S, KeyEvent.CTRL_DOWN_MASK));
        exportItem.addActionListener(this::onExport);

        JMenuItem clearItem = new JMenuItem("Clear View");
        clearItem.addActionListener(e -> clearView());

        JMenuItem exitItem = new JMenuItem("Exit");
        exitItem.addActionListener(e -> dispose());

        file.add(importItem);
        file.add(exportItem);
        file.addSeparator();
        file.add(clearItem);
        file.add(exitItem);

        JMenu dataMenu = new JMenu("Data");

        JMenuItem findItem = new JMenuItem("Find Next");
        findItem.setAccelerator(KeyStroke.getKeyStroke(KeyEvent.VK_F, KeyEvent.CTRL_DOWN_MASK));
        findItem.addActionListener(this::onFindNext);

        JMenuItem gotoRowItem = new JMenuItem("Go To Row");
        gotoRowItem.addActionListener(this::onGotoRow);

        JMenuItem deleteRowItem = new JMenuItem("Delete Selected Row(s)");
        deleteRowItem.setAccelerator(KeyStroke.getKeyStroke(KeyEvent.VK_DELETE, 0));
        deleteRowItem.addActionListener(this::onDeleteSelectedRows);

        JMenuItem undoDeleteItem = new JMenuItem("Undo Delete");
        undoDeleteItem.setAccelerator(KeyStroke.getKeyStroke(KeyEvent.VK_Z, KeyEvent.CTRL_DOWN_MASK));
        undoDeleteItem.addActionListener(this::onUndoDelete);

        JCheckBoxMenuItem lockEditItem = new JCheckBoxMenuItem("Lock Editing");
        lockEditItem.addActionListener(e -> {
            editLocked = lockEditItem.isSelected();
            applyEditLock();
        });

        dataMenu.add(findItem);
        dataMenu.add(gotoRowItem);
        dataMenu.addSeparator();
        dataMenu.add(deleteRowItem);
        dataMenu.add(undoDeleteItem);
        dataMenu.addSeparator();
        dataMenu.add(lockEditItem);

        bar.add(file);
        bar.add(dataMenu);
        return bar;
    }

    private void onImport(ActionEvent e) {
        JFileChooser chooser = new JFileChooser();
        chooser.setDialogTitle("Select an Excel file");
        chooser.setFileFilter(new FileNameExtensionFilter("Excel Workbook (*.xlsx)", "xlsx"));

        if (chooser.showOpenDialog(this) != JFileChooser.APPROVE_OPTION) return;

        try {
            currentFilePath = chooser.getSelectedFile().getAbsolutePath();
            currentData = importer.load(currentFilePath);
            buildEditableGrid();

            fileLabel.setText("Loaded: " + chooser.getSelectedFile().getName());
            statusLabel.setText("Imported successfully. You can edit, search, jump to row, and delete rows.");

            undoStack.clear();
            resetSearchState();

        } catch (Exception ex) {
            JOptionPane.showMessageDialog(this,
                    "Import failed: " + ex.getMessage(),
                    "Import Error",
                    JOptionPane.ERROR_MESSAGE);
        }
    }

    private void buildEditableGrid() {
        if (currentData == null) return;

        int colCount = currentData.columnCount;
        int dataRows = Math.max(0, currentData.rowCount - currentData.headerRows);

        String[] columns = extractColumns(currentData);
        Object[][] dataRowsMatrix = new Object[dataRows][colCount];

        for (int r = 0; r < dataRows; r++) {
            for (int c = 0; c < colCount; c++) {
                String value = currentData.cells[currentData.headerRows + r][c];
                dataRowsMatrix[r][c] = value != null ? value : "";
            }
        }

        suppressSync = true;
        tableModel = new EditableTableModel(dataRowsMatrix, columns);
        dataTable = new JTable(tableModel);
        sorter = new TableRowSorter<EditableTableModel>(tableModel);
        dataTable.setRowSorter(sorter);

        configureTableLook();
        applyEditLock();

        tableModel.addTableModelListener(evt -> {
            if (suppressSync) return;
            if (evt.getType() == TableModelEvent.UPDATE) {
                syncGridToCurrentData();
            }
        });

        JScrollPane scroll = new JScrollPane(dataTable);
        scroll.setBorder(BorderFactory.createLineBorder(new Color(214, 223, 236)));

        tableHostPanel.removeAll();
        tableHostPanel.add(scroll, BorderLayout.CENTER);
        tableHostPanel.revalidate();
        tableHostPanel.repaint();

        suppressSync = false;
    }

    private void configureTableLook() {
        dataTable.setFont(new Font("Segoe UI", Font.PLAIN, 13));
        dataTable.setRowHeight(28);
        dataTable.setShowHorizontalLines(true);
        dataTable.setShowVerticalLines(true);
        dataTable.setGridColor(new Color(229, 235, 244));
        dataTable.setSelectionBackground(new Color(32, 108, 172));
        dataTable.setSelectionForeground(Color.WHITE);
        dataTable.setAutoResizeMode(JTable.AUTO_RESIZE_OFF);

        JTableHeader header = dataTable.getTableHeader();
        header.setFont(new Font("Segoe UI", Font.BOLD, 13));
        header.setBackground(new Color(25, 74, 120));
        header.setForeground(Color.WHITE);
        header.setPreferredSize(new Dimension(header.getWidth(), 32));
        header.setReorderingAllowed(false);
        header.setDefaultRenderer(new DefaultTableCellRenderer() {
            private static final long serialVersionUID = 1L;

            {
                setHorizontalAlignment(SwingConstants.CENTER);
                setOpaque(true);
            }

            @Override
            public java.awt.Component getTableCellRendererComponent(
                    JTable table,
                    Object value,
                    boolean isSelected,
                    boolean hasFocus,
                    int row,
                    int column) {
                super.getTableCellRendererComponent(table, value, isSelected, hasFocus, row, column);
                setBackground(new Color(25, 74, 120));
                setForeground(Color.WHITE);
                setFont(new Font("Segoe UI", Font.BOLD, 13));
                setBorder(BorderFactory.createCompoundBorder(
                        BorderFactory.createMatteBorder(0, 0, 1, 1, new Color(18, 56, 91)),
                        BorderFactory.createEmptyBorder(6, 8, 6, 8)));

                String text = value != null ? value.toString().trim() : "";
                setText(text.isEmpty() ? ("Column " + (column + 1)) : text);
                return this;
            }
        });

        for (int c = 0; c < dataTable.getColumnCount(); c++) {
            dataTable.getColumnModel().getColumn(c).setPreferredWidth(180);
        }

        dataTable.setDefaultRenderer(Object.class, new DefaultTableCellRenderer() {
            private static final long serialVersionUID = 1L;

            @Override
            public java.awt.Component getTableCellRendererComponent(
                    JTable table,
                    Object value,
                    boolean isSelected,
                    boolean hasFocus,
                    int row,
                    int column) {
                java.awt.Component cell = super.getTableCellRendererComponent(
                        table, value, isSelected, hasFocus, row, column);
                if (!isSelected) {
                    cell.setBackground((row % 2 == 0) ? Color.WHITE : new Color(235, 246, 255));
                }
                return cell;
            }
        });
    }

    private String[] extractColumns(ExcelData data) {
        String[] columns = new String[data.columnCount];
        int headerRow = Math.max(0, Math.min(data.headerRows - 1, data.rowCount - 1));

        Map<String, Integer> seen = new HashMap<String, Integer>();
        for (int c = 0; c < data.columnCount; c++) {
            String base = data.cells[headerRow][c];
            if (base == null || base.trim().isEmpty()) {
                base = "Column " + (c + 1);
            } else {
                base = base.trim();
            }

            int count = seen.containsKey(base) ? seen.get(base) + 1 : 1;
            seen.put(base, count);
            columns[c] = (count == 1) ? base : (base + " (" + count + ")");
        }
        return columns;
    }

    private void onExport(ActionEvent e) {
        if (currentData == null) {
            JOptionPane.showMessageDialog(this, "Nothing to export. Import a file first.");
            return;
        }

        if (dataTable != null && dataTable.isEditing()) {
            dataTable.getCellEditor().stopCellEditing();
        }
        syncGridToCurrentData();

        JFileChooser chooser = new JFileChooser();
        chooser.setDialogTitle("Export Excel file");
        chooser.setSelectedFile(new java.io.File(defaultExportName()));
        chooser.setFileFilter(new FileNameExtensionFilter("Excel Workbook (*.xlsx)", "xlsx"));

        if (chooser.showSaveDialog(this) != JFileChooser.APPROVE_OPTION) return;

        try {
            String outputPath = chooser.getSelectedFile().getAbsolutePath();
            exporter.export(currentData, outputPath);
            JOptionPane.showMessageDialog(this, "Export successful.");
            statusLabel.setText("Exported: " + outputPath);
        } catch (Exception ex) {
            JOptionPane.showMessageDialog(this,
                    "Export failed: " + ex.getMessage(),
                    "Export Error",
                    JOptionPane.ERROR_MESSAGE);
        }
    }

    private void onFindNext(ActionEvent e) {
        if (tableModel == null) {
            JOptionPane.showMessageDialog(this, "Import a file first.");
            return;
        }

        String query = searchField.getText().trim();
        if (query.isEmpty()) {
            JOptionPane.showMessageDialog(this, "Enter a value to search.");
            searchField.requestFocusInWindow();
            return;
        }

        if (!query.equalsIgnoreCase(lastQuery)) {
            lastQuery = query;
            lastFoundModelRow = -1;
            lastFoundModelCol = -1;
        }

        int[] match = findNextCell(query);
        if (match == null) {
            JOptionPane.showMessageDialog(this, "No match found for: " + query);
            resetSearchState();
            return;
        }

        focusCell(match[0], match[1]);
        lastFoundModelRow = match[0];
        lastFoundModelCol = match[1];
        statusLabel.setText("Match found at row " + (match[0] + 1) + ", column " + (match[1] + 1));
    }

    private int[] findNextCell(String query) {
        String needle = query.toLowerCase(Locale.ROOT);
        int rows = tableModel.getRowCount();
        int cols = tableModel.getColumnCount();
        if (rows == 0 || cols == 0) return null;

        int total = rows * cols;
        int start = 0;
        if (lastFoundModelRow >= 0 && lastFoundModelCol >= 0) {
            start = lastFoundModelRow * cols + lastFoundModelCol + 1;
        }

        for (int offset = 0; offset < total; offset++) {
            int idx = (start + offset) % total;
            int row = idx / cols;
            int col = idx % cols;
            String value = Objects.toString(tableModel.getValueAt(row, col), "").toLowerCase(Locale.ROOT);
            if (value.contains(needle)) return new int[]{row, col};
        }
        return null;
    }

    private void onGotoRow(ActionEvent e) {
        if (tableModel == null) {
            JOptionPane.showMessageDialog(this, "Import a file first.");
            return;
        }

        String input = JOptionPane.showInputDialog(this, "Go to row (1 to " + tableModel.getRowCount() + "):");
        if (input == null) return;

        try {
            int row = Integer.parseInt(input.trim());
            if (row < 1 || row > tableModel.getRowCount()) {
                JOptionPane.showMessageDialog(this, "Row out of range.");
                return;
            }
            focusCell(row - 1, 0);
            statusLabel.setText("Moved to row " + row + ".");
        } catch (NumberFormatException ex) {
            JOptionPane.showMessageDialog(this, "Enter a valid row number.");
        }
    }

    private void onDeleteSelectedRows(ActionEvent e) {
        if (tableModel == null || dataTable == null) {
            JOptionPane.showMessageDialog(this, "Import a file first.");
            return;
        }

        int[] selectedViewRows = dataTable.getSelectedRows();
        if (selectedViewRows.length == 0) {
            JOptionPane.showMessageDialog(this, "Select at least one row to delete.");
            return;
        }

        int confirm = JOptionPane.showConfirmDialog(this,
                "Delete " + selectedViewRows.length + " selected row(s)?",
                "Confirm Delete",
                JOptionPane.YES_NO_OPTION);
        if (confirm != JOptionPane.YES_OPTION) return;

        Integer[] modelRows = new Integer[selectedViewRows.length];
        for (int i = 0; i < selectedViewRows.length; i++) {
            modelRows[i] = dataTable.convertRowIndexToModel(selectedViewRows[i]);
        }
        Arrays.sort(modelRows, Comparator.<Integer>naturalOrder());

        List<Integer> removedAt = new ArrayList<Integer>();
        List<Object[]> removedRows = new ArrayList<Object[]>();

        for (int modelRow : modelRows) {
            removedAt.add(modelRow);
            Object[] rowData = new Object[tableModel.getColumnCount()];
            for (int c = 0; c < tableModel.getColumnCount(); c++) {
                rowData[c] = tableModel.getValueAt(modelRow, c);
            }
            removedRows.add(rowData);
        }

        suppressSync = true;
        for (int i = modelRows.length - 1; i >= 0; i--) {
            tableModel.removeRow(modelRows[i]);
        }
        suppressSync = false;

        undoStack.push(new DeletedRowsSnapshot(removedAt, removedRows));
        syncGridToCurrentData();
        statusLabel.setText("Deleted " + selectedViewRows.length + " row(s). Use Undo Delete to restore.");
    }

    private void onUndoDelete(ActionEvent e) {
        if (tableModel == null) {
            JOptionPane.showMessageDialog(this, "Nothing loaded.");
            return;
        }
        if (undoStack.isEmpty()) {
            JOptionPane.showMessageDialog(this, "No deleted rows to restore.");
            return;
        }

        DeletedRowsSnapshot snapshot = undoStack.pop();

        suppressSync = true;
        for (int i = 0; i < snapshot.modelRows.size(); i++) {
            int idx = Math.min(snapshot.modelRows.get(i).intValue(), tableModel.getRowCount());
            tableModel.insertRow(idx, snapshot.rows.get(i));
        }
        suppressSync = false;

        syncGridToCurrentData();
        statusLabel.setText("Undo complete. Deleted rows restored.");
    }

    private void onToggleEditLock(ActionEvent e) {
        editLocked = !editLocked;
        applyEditLock();
    }

    private void applyEditLock() {
        if (tableModel != null) {
            tableModel.setEditable(!editLocked);
        }
        statusLabel.setText(editLocked ? "Editing locked." : "Editing unlocked.");
    }

    private void focusCell(int modelRow, int modelCol) {
        if (dataTable == null) return;

        int viewRow = dataTable.convertRowIndexToView(modelRow);
        if (viewRow < 0) return;

        dataTable.changeSelection(viewRow, modelCol, false, false);
        Rectangle rect = dataTable.getCellRect(viewRow, modelCol, true);
        dataTable.scrollRectToVisible(rect);
    }

    private void syncGridToCurrentData() {
        if (currentData == null || tableModel == null) return;

        int headerRows = currentData.headerRows;
        int columnCount = currentData.columnCount;
        int totalRows = headerRows + tableModel.getRowCount();

        String[][] rebuilt = new String[totalRows][columnCount];

        for (int r = 0; r < Math.min(headerRows, currentData.rowCount); r++) {
            for (int c = 0; c < columnCount; c++) {
                String value = currentData.cells[r][c];
                rebuilt[r][c] = value != null ? value : "";
            }
        }

        for (int r = 0; r < tableModel.getRowCount(); r++) {
            for (int c = 0; c < columnCount; c++) {
                Object value = tableModel.getValueAt(r, c);
                rebuilt[headerRows + r][c] = value != null ? value.toString() : "";
            }
        }

        List<CellRangeAddress> headerMergedRegions = new ArrayList<CellRangeAddress>();
        for (CellRangeAddress region : currentData.mergedRegions) {
            if (region.getLastRow() < headerRows && region.getLastColumn() < columnCount) {
                headerMergedRegions.add(new CellRangeAddress(
                        region.getFirstRow(),
                        region.getLastRow(),
                        region.getFirstColumn(),
                        region.getLastColumn()));
            }
        }

        currentData = new ExcelData(rebuilt, headerRows, columnCount, totalRows, headerMergedRegions);
    }

    private String defaultExportName() {
        if (currentFilePath == null || currentFilePath.isEmpty()) return "exported.xlsx";
        String fileName = new java.io.File(currentFilePath).getName();
        int idx = fileName.lastIndexOf('.');
        if (idx > 0) fileName = fileName.substring(0, idx);
        return fileName + "_edited_export.xlsx";
    }

    private void resetSearchState() {
        lastFoundModelRow = -1;
        lastFoundModelCol = -1;
        lastQuery = "";
    }

    private void clearView() {
        currentData = null;
        currentFilePath = null;
        tableModel = null;
        dataTable = null;
        sorter = null;
        undoStack.clear();
        resetSearchState();
        fileLabel.setText("No file loaded");

        tableHostPanel.removeAll();
        tableHostPanel.add(buildEmptyStatePanel(), BorderLayout.CENTER);
        tableHostPanel.revalidate();
        tableHostPanel.repaint();

        statusLabel.setText("Cleared. Import an Excel file to begin.");
    }

    private static class EditableTableModel extends DefaultTableModel {
        private static final long serialVersionUID = 1L;

        private boolean editable = true;

        EditableTableModel(Object[][] data, Object[] columnNames) {
            super(data, columnNames);
        }

        void setEditable(boolean editable) {
            this.editable = editable;
        }

        @Override
        public boolean isCellEditable(int row, int column) {
            return editable;
        }
    }

    private static class DeletedRowsSnapshot {
        private final List<Integer> modelRows;
        private final List<Object[]> rows;

        DeletedRowsSnapshot(List<Integer> modelRows, List<Object[]> rows) {
            this.modelRows = modelRows;
            this.rows = rows;
        }
    }
}
