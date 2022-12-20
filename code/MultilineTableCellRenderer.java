package net.codejava;

import java.awt.Color;
import java.awt.Component;
import java.awt.Dimension;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import javax.swing.JLabel;
import javax.swing.JTable;
import javax.swing.JTextArea;
import javax.swing.UIManager;
import javax.swing.border.EmptyBorder;
import javax.swing.table.DefaultTableCellRenderer;
import javax.swing.table.TableCellRenderer;
import javax.swing.text.StyledDocument;

public class MultilineTableCellRenderer extends JTextArea
implements TableCellRenderer {
    private List<List<Integer>> rowColHeight = new ArrayList<>();
    private static final int MAX_LEN = 19500;
    public static final DefaultTableCellRenderer CENTER_ALIGNMENT = new DefaultTableCellRenderer();
    Map<String, String> states = new HashMap<String, String>();
    
    public MultilineTableCellRenderer() {
        setLineWrap(true);
        setWrapStyleWord(true);
        setOpaque(true);
    }
    
    protected String shortener(String str) {
        if (str.length() < MAX_LEN) {       	
            return str;
        } else {
        	return str.substring(0, MAX_LEN - 10) + "...";
        }
    }
    
    @Override
    public Component getTableCellRendererComponent(
            JTable table, Object value, boolean isSelected, boolean hasFocus,
            int row, int column) {
    	Component renderer = CENTER_ALIGNMENT.getTableCellRendererComponent(table, value, isSelected, hasFocus, row, column);
    	
        if (isSelected) {
            setForeground(table.getSelectionForeground());
            setBackground(table.getSelectionBackground());
        } else {
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
        } else {
            setBorder(new EmptyBorder(1, 2, 1, 2));
        }
        
        if (value != null) {
            setText(shortener(value.toString()));
        } else {
            setText("");
        }
        setForeground( Color.black );
        setBackground( Color.white );
        
    	adjustRowHeight( table, row, column );
    	
    	if ( states.get( table.getValueAt(row, column).toString() ) == "true" ) { 
	        // ((JLabel) renderer).setOpaque(false);
		    ((JLabel) renderer).setHorizontalAlignment(JLabel.CENTER);
		    ((JLabel) renderer).setVerticalAlignment(JLabel.CENTER);
	    	return renderer;
	    } else {	         	    
	    	return this; 
	    }
    }

    private void adjustRowHeight(JTable table, int row, int column) {
        int cWidth = table.getTableHeader().getColumnModel().getColumn(column).getWidth();
        setSize(new Dimension(cWidth, 1000));
        int prefH = getPreferredSize().height;
        
        while (rowColHeight.size() <= row) {
            rowColHeight.add(new ArrayList<Integer>(column));
        }
        List<Integer> colHeights = rowColHeight.get(row);
        
        while (colHeights.size() <= column) {
            colHeights.add(0);
        }
        
        colHeights.set(column, prefH);
        int maxH = prefH;
        for (Integer colHeight : colHeights) {
            if (colHeight > maxH) {
                maxH = colHeight;
            }
        }
        
        if ( table.getRowHeight(row) <= maxH ) {
        	if ( table.getRowHeight(row) != maxH ) {       		
        	 	states.put( table.getValueAt(row, column).toString(), "true" );
        	} else {       		       		
        		if ( maxH == 18 ) {       		
            		states.put( table.getValueAt(row, column).toString(), "true" );
            	} else {
            		// System.out.println( table.getRowHeight(row) + " " + table.getValueAt(row, column) + " " + maxH );
            		states.put( table.getValueAt(row, column).toString(), "false" );
            	}
        	}
            table.setRowHeight( row, maxH );                   	
        } else {       	       	
        	if ( maxH == 18 ) {       		
        		states.put( table.getValueAt(row, column).toString(), "true" );
        	} else {
        		// System.out.println( table.getRowHeight(row) + " " + table.getValueAt(row, column) + " " + maxH );
        		states.put( table.getValueAt(row, column).toString(), "false" );
        	}
        }
    }
}
