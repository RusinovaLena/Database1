package net.codejava;

import javax.swing.*;
import javax.swing.table.JTableHeader;
import javax.swing.table.TableCellRenderer;
import javax.swing.table.TableColumn;
import javax.swing.table.TableColumnModel;
import java.awt.*;
import java.util.Enumeration;
import java.util.Vector;


public class GroupableTableHeader extends JTableHeader {
    private static final String uiClassID = "GroupableTableHeaderUI";
    private Vector<ColumnGroup> columnGroups = null;

    public GroupableTableHeader(TableColumnModel model) {
        super(model);
        setUI(new GroupableTableHeaderUI());
        setReorderingAllowed(false);
    }
    public void updateUI(){

        TableCellRenderer renderer = getDefaultRenderer();
        if (renderer instanceof Component) {
            SwingUtilities.updateComponentTreeUI((Component)renderer);
        }
    }

    public void setReorderingAllowed(boolean b) {
        reorderingAllowed = false;
    }

    public void addColumnGroup(ColumnGroup g) {
        if (columnGroups == null) {
            columnGroups = new Vector();
        }
        columnGroups.addElement(g);
    }

    public Enumeration getColumnGroups(TableColumn col) {
        if (columnGroups == null) return null;
        Enumeration<ColumnGroup> e = columnGroups.elements();
        while (e.hasMoreElements()) {
            ColumnGroup cGroup = e.nextElement();
            Vector v_ret = cGroup.getColumnGroups(col,new Vector());
            if (v_ret != null) {
                return v_ret.elements();
            }
        }
        return null;
    }

    public void setColumnMargin() {
        if (columnGroups == null) return;
        int columnMargin = getColumnModel().getColumnMargin();
        Enumeration<ColumnGroup> e = columnGroups.elements();
        while (e.hasMoreElements()) {
            ColumnGroup cGroup = e.nextElement();
            cGroup.setColumnMargin(columnMargin);
        }
    }
}

