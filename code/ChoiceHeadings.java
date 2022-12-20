package net.codejava;

import java.awt.Color;
import java.awt.Component;
import java.awt.Font;
import java.awt.GridLayout;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.IOException;
import java.util.ArrayList;

import javax.swing.DefaultComboBoxModel;
import javax.swing.JCheckBox;
import javax.swing.JComboBox;
import javax.swing.JList;
import javax.swing.JPanel;
import javax.swing.ListCellRenderer;

import org.apache.poi.ss.usermodel.Sheet;


public class ChoiceHeadings implements ActionListener {
	
	static CustomerItem[] headersString;	
	
	public class CustomerItem {

        public String label;
        public boolean status;

        public CustomerItem(String label, boolean status) {
            this.label = label;
            this.status = status;
        }
    }
    
    public class RenderCheckComboBox implements ListCellRenderer {

        //a JCheckBox is associated for one item
        JCheckBox checkBox;

        Color selectedBG = new Color(112, 146, 190);

        public RenderCheckComboBox() {
            this.checkBox = new JCheckBox();
        }

        @Override
        public Component getListCellRendererComponent(JList list, Object value, int index, boolean isSelected,
                boolean cellHasFocus) {

            //recuperate the item value
            CustomerItem value_ = (CustomerItem) value;

            if (value_ != null) {
                //put the label of item as a label for the associated JCheckBox object
                checkBox.setText(value_.label);

                //put the status of item as a status for the associated JCheckBox object
                checkBox.setSelected(value_.status);
            }

            if (isSelected) {
                checkBox.setBackground(Color.GRAY);
            } else {
                checkBox.setBackground(Color.WHITE);
            }
            return checkBox;
        }
    }
    
    // ��������� JCheckBox 
    public void actionPerformed(ActionEvent e) {
        JComboBox cb = (JComboBox) e.getSource();
        CheckComboStore store = (CheckComboStore)cb.getSelectedItem();
        CheckComboRenderer ccr = (CheckComboRenderer)cb.getRenderer();
        ccr.checkBox.setSelected((store.state = !store.state));
    }
	
	public JComboBox<CustomerItem> outputPanel(ArrayList headers, String firstString) {

		 JComboBox<CustomerItem> combo = new JComboBox<CustomerItem>() {
	            @Override
	            public void setPopupVisible(boolean visible) {
	                if (visible) {
	                    super.setPopupVisible(visible);
	                }
	            }
	        };
	        
	        headersString = new CustomerItem[headers.size() + 1];
	        headersString[0] = new CustomerItem(firstString, true);
	        for (int i = 0; i < headers.size(); i++) {
	        	headersString[i + 1] = new CustomerItem(headers.get(i).toString(), false);
	        }	
	        
	        combo.setModel(new DefaultComboBoxModel<CustomerItem>(headersString));
	        combo.setRenderer(new RenderCheckComboBox());
	        
	        combo.addActionListener(new ActionListener() {
	            @Override
	            public void actionPerformed(ActionEvent ae) {
	                CustomerItem item = (CustomerItem) ((JComboBox) ae.getSource()).getSelectedItem();
	                item.status = !item.status;

	                // update the ui of combo
	                combo.updateUI();

	                //keep the popMenu of the combo as visible
	                combo.setPopupVisible(true);
	                for (int i = 1; i < headersString.length; i++) {
	                	if (headersString[i].status == true) {
	                		headersString[0].status = false;
	                	}
	                }	
	            }
	        });
	        
	        return combo;
	}
	
	public ArrayList selectedHeaders() {
		
		ArrayList<Integer> headers = new ArrayList<Integer>();
			
    	for (int i = 0; i < headersString.length; i++) {
    		if (headersString[i].status == true) {
    			headers.add(i);
    		}	
    	}
    	
    	return headers;
	}
	
	public ArrayList<Integer> searchRows(ArrayList headers, Sheet currentSheet) {
		
		ArrayList<Integer> headersString = selectedHeaders();
	    ArrayList<Integer> currentRows = new ArrayList<Integer>();
	    
	    int currentZ[] = new int[2];
	    try {
	    	if (headersString.get(0) == 0) {
	    		currentRows.add(0);
            	currentRows.add(currentSheet.getLastRowNum() + 1);
	    	} else {
		    	if (headersString.size() == 1) {
	            	currentZ = new SearchHeaders().searchTwoRow(headers.get(headersString.get(0) - 1).toString());
	            	currentRows.add(currentZ[0]);
	            	currentRows.add(currentZ[1]);
	
	        	} else {	
	            	for (int i = 0; i < headersString.size();i++) {
	            		currentZ = new SearchHeaders().searchTwoRow(headers.get(headersString.get(i) - 1).toString());
	            		currentRows.add(currentZ[0]);
	                	currentRows.add(currentZ[1]);
	            	}
	        	}
	    	}	
			
		} catch (IOException e2) {
			e2.printStackTrace();
		}
	    	  	    
	    return currentRows;
	}
}
