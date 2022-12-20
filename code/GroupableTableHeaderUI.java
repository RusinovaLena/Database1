package net.codejava;

import javax.swing.*;
import javax.swing.event.MouseInputListener;
import javax.swing.plaf.basic.BasicTableHeaderUI;
import javax.swing.table.DefaultTableCellRenderer;
import javax.swing.table.TableCellRenderer;
import javax.swing.table.TableColumn;
import javax.swing.table.TableColumnModel;
import java.awt.*;
import java.util.Enumeration;
import java.util.Hashtable;
import java.util.Vector;



public class GroupableTableHeaderUI extends BasicTableHeaderUI {

    @Override
    protected MouseInputListener createMouseInputListener() {
        return super.createMouseInputListener();
    }

    public void paint(Graphics g, JComponent c) {
        Rectangle clipBounds = g.getClipBounds();
        if (header.getColumnModel() == null) return;
        ((GroupableTableHeader)header).setColumnMargin();
        int column = 0;
        Dimension size = header.getSize();
        Rectangle cellRect  = new Rectangle(0, 0, size.width, size.height);
        Hashtable<ColumnGroup, Rectangle> h = new Hashtable();
        int columnMargin = header.getColumnModel().getColumnMargin();

        Enumeration<TableColumn> enumeration = header.getColumnModel().getColumns();
        Vector<ColumnGroup> painted = new Vector();
        while (enumeration.hasMoreElements()) {
            cellRect.height = size.height;
            cellRect.y      = 0;
            TableColumn aColumn = enumeration.nextElement();
            Enumeration cGroups = ((GroupableTableHeader)header).getColumnGroups(aColumn);
            if (cGroups != null) {
                int groupHeight = 0;
                while (cGroups.hasMoreElements()) {
                    ColumnGroup cGroup = (ColumnGroup)cGroups.nextElement();
                    Rectangle groupRect = h.get(cGroup);
                    if (groupRect == null) {
                        groupRect = new Rectangle(cellRect);
                        Dimension d = cGroup.getSize(header.getTable());
                        groupRect.width  = d.width;
                        groupRect.height = d.height;
                        h.put(cGroup, groupRect);
                    }
                    if(!painted.contains(cGroup))
                    {
                        paintCell(g, groupRect, cGroup);
                        painted.addElement(cGroup);
                    }
                    groupHeight += groupRect.height;
                    cellRect.height = size.height - groupHeight;
                    cellRect.y      = groupHeight;
                }
            }
            cellRect.width = aColumn.getWidth();
            if (cellRect.intersects(clipBounds)) {
                paintCell(g, cellRect, column);
            }
            cellRect.x += cellRect.width;
            column++;
        }
    }

    private void paintCell(Graphics g, Rectangle cellRect, int columnIndex) {
        TableColumn aColumn = header.getColumnModel().getColumn(columnIndex);
        TableCellRenderer renderer = aColumn.getHeaderRenderer();
        //revised by Java2s.com
        renderer = new DefaultTableCellRenderer(){
            public Component getTableCellRendererComponent(JTable table, Object value, boolean isSelected, boolean hasFocus, int row, int column) {
                JLabel header = new JLabel();
                header.setForeground(table.getTableHeader().getForeground());
                header.setBackground(table.getTableHeader().getBackground());
                header.setFont(table.getTableHeader().getFont());

                header.setHorizontalAlignment(JLabel.CENTER);
                header.setText(value.toString());
                header.setBorder(UIManager.getBorder("TableHeader.cellBorder"));
                return header;
            }

        };
        Component c = renderer.getTableCellRendererComponent(
                header.getTable(), aColumn.getHeaderValue(),false, false, -1, columnIndex);

        c.setBackground(UIManager.getColor("control"));

        rendererPane.add(c);
        rendererPane.paintComponent(g, c, header, cellRect.x, cellRect.y,
                cellRect.width, cellRect.height, true);
    }

    private void paintCell(Graphics g, Rectangle cellRect, ColumnGroup cGroup) {
        TableCellRenderer renderer = cGroup.getHeaderRenderer();

        Component component = renderer.getTableCellRendererComponent(
                header.getTable(), cGroup.getHeaderValue(),false, false, -1, -1);
        rendererPane.add(component);
        rendererPane.paintComponent(g, component, header, cellRect.x, cellRect.y,
                cellRect.width, cellRect.height, true);
    }

    private int getHeaderHeight() {
        int height = 0;
        TableColumnModel columnModel = header.getColumnModel();
        for(int column = 0; column < columnModel.getColumnCount(); column++) {
            TableColumn aColumn = columnModel.getColumn(column);
            TableCellRenderer renderer = aColumn.getHeaderRenderer();
            if(renderer == null){
                return 60;
            }

            Component comp = renderer.getTableCellRendererComponent(
                    header.getTable(), aColumn.getHeaderValue(), false, false,-1, column);
            int cHeight = comp.getPreferredSize().height;
            Enumeration e = ((GroupableTableHeader)header).getColumnGroups(aColumn);
            if (e != null) {
                while (e.hasMoreElements()) {
                    ColumnGroup cGroup = (ColumnGroup)e.nextElement();
                    cHeight += cGroup.getSize(header.getTable()).height;
                }
            }
            height = Math.max(height, cHeight);
        }
        return height;
    }

    private Dimension createHeaderSize(long width) {
        TableColumnModel columnModel = header.getColumnModel();
        width += columnModel.getColumnMargin() * columnModel.getColumnCount();
        if (width > Integer.MAX_VALUE) {
            width = Integer.MAX_VALUE;
        }
        return new Dimension((int)width, getHeaderHeight());
    }

    public Dimension getPreferredSize(JComponent c) {
        long width = 0;
        Enumeration<TableColumn> enumeration = header.getColumnModel().getColumns();
        while (enumeration.hasMoreElements()) {
            TableColumn aColumn = enumeration.nextElement();
            width = width + aColumn.getPreferredWidth();
        }
        return createHeaderSize(width);
    }
}