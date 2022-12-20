package net.codejava;

import javax.swing.*;
import javax.swing.table.DefaultTableCellRenderer;
import javax.swing.table.JTableHeader;
import javax.swing.table.TableCellRenderer;
import javax.swing.table.TableColumn;
import java.awt.*;
import java.util.Enumeration;
import java.util.Vector;


public class ColumnGroup {
    private TableCellRenderer renderer;
    private Vector<Object> v;
    protected String text;
    protected int margin=0;

    public ColumnGroup(String text) {
        this(null,text);
    }

    private ColumnGroup(TableCellRenderer renderer, String text) {
        if (renderer == null) {
            this.renderer = new DefaultTableCellRenderer() {
                public Component getTableCellRendererComponent(JTable table, Object value,
                                                               boolean isSelected, boolean hasFocus, int row, int column) {
                    JTableHeader header = table.getTableHeader();
                    if (header != null) {
                        setForeground(header.getForeground());
                        setBackground(header.getBackground());
                        setFont(header.getFont());
                    }
                    setHorizontalAlignment(JLabel.CENTER);
                    setText((value == null) ? "" : value.toString());
                    setBorder(UIManager.getBorder("TableHeader.cellBorder"));
                    return this;
                }
            };
        } else {
            this.renderer = renderer;
        }
        this.text = text;
        v = new Vector();
    }

    public void add(Object obj) {
        if (obj == null) { return; }
        v.addElement(obj);
    }

    Vector getColumnGroups(TableColumn c, Vector g) {
        g.addElement(this);
        if (v.contains(c)) return g;
        Enumeration<Object> e = v.elements();
        while (e.hasMoreElements()) {
            Object obj = e.nextElement();
            if (obj instanceof ColumnGroup) {
                Vector groups =
                        ((ColumnGroup)obj).getColumnGroups(c,(Vector)g.clone());
                if (groups != null) return groups;
            }
        }
        return null;
    }

    TableCellRenderer getHeaderRenderer() {
        return renderer;
    }

    public void setHeaderRenderer(TableCellRenderer renderer) {
        if (renderer != null) {
            this.renderer = renderer;
        }
    }

    Object getHeaderValue() {
        return text;
    }

    private static boolean contains(Enumeration enumeration, Object o) {

        while (enumeration.hasMoreElements())
            if(enumeration.nextElement() == o)
                return true;

        return false;
    }

    Dimension getSize(JTable table) {
        Component comp = renderer.getTableCellRendererComponent(
                table, getHeaderValue(), false, false,-1, -1);
        int height = comp.getPreferredSize().height;
        int width  = 0;
        Enumeration<Object> en = v.elements();
        while (en.hasMoreElements()) {
            Object obj = en.nextElement();
            if (obj instanceof TableColumn) {
                if(contains(table.getColumnModel().getColumns(), obj)){
                    TableColumn aColumn = (TableColumn) obj;
                    width += aColumn.getWidth();
//                  width += margin;
                }
            } else {
                width += ((ColumnGroup)obj).getSize(table).width;
            }
        }
        return new Dimension(width, height);
    }

    void setColumnMargin(int margin) {
        this.margin = margin;
        Enumeration<Object> e = v.elements();
        while (e.hasMoreElements()) {
            Object obj = e.nextElement();
            if (obj instanceof ColumnGroup) {
                ((ColumnGroup)obj).setColumnMargin(margin);
            }
        }
    }
}