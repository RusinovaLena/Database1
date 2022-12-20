package net.codejava;

import java.awt.Component;

import javax.swing.JTable;
import javax.swing.JTextArea;
import javax.swing.UIManager;
import javax.swing.border.EmptyBorder;
import javax.swing.table.AbstractTableModel;
import javax.swing.table.TableCellRenderer;

class MyTableModel extends AbstractTableModel{
private String[][] data;
private String[] columnNames;

    public MyTableModel(String[][] d, String[] cn) {
        data = d;
        columnNames = cn;
    }
    public boolean isCellEditable(int row, int col) {
        return false;
    }
 
    public Object getValueAt(int row, int col) {
        return data[row][col];
    }
 
    public int getColumnCount() {
        return columnNames.length;
    }
 
    public int getRowCount() {
        return data.length;
    }
 
    public String getColumnName(int col) {
        return columnNames[col];
    }
}
 
class MultiLineCellRenderer extends JTextArea implements TableCellRenderer {
 
    public MultiLineCellRenderer() {
        setLineWrap(true);
        setWrapStyleWord(true);
        setOpaque(true);
        setEditable(false); //this line doesn't seem to be doing anything
    }
 
    public Component getTableCellRendererComponent(JTable table, Object value,
        boolean isSelected, boolean hasFocus, int row, int column) {
        if (isSelected) {
            setForeground(table.getSelectionForeground());
            setBackground(table.getSelectionBackground());
        }
        else {
            setForeground(table.getForeground());
            setBackground(table.getBackground());
        }
        setFont(table.getFont());
        if (hasFocus) {
            setBorder(UIManager.getBorder("Table.focusCellHighlightBorder"));
            if (table.isCellEditable(row, column)) {
                setForeground(UIManager.getColor("Table.focusCellForeground"));
                setBackground(UIManager.getColor("Table.focusCellBackground"));
            }
        }
        else {
            setBorder(new EmptyBorder(1, 2, 1, 2));
        }
        setText((value == null) ? "" : value.toString());
        setEditable(false); //this line doesn't seem to be doing anything
        return this;
    }
}
