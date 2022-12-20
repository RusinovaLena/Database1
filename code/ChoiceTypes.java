package net.codejava;

import java.awt.Color;
import java.awt.Component;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.util.ArrayList;

import javax.swing.DefaultComboBoxModel;
import javax.swing.JCheckBox;
import javax.swing.JComboBox;
import javax.swing.JList;
import javax.swing.ListCellRenderer;

import net.codejava.ChoiceTypes.CustomerItem;
import net.codejava.ChoiceTypes.RenderCheckComboBox;

public class ChoiceTypes implements ActionListener {
	
	static CustomerItem[] types;	
	
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
    
    // несколько JCheckBox 
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
	        
	        types = new CustomerItem[headers.size() + 1];
	        types[0] = new CustomerItem(firstString, true);
	        for (int i = 0; i < headers.size(); i++) {
	        	types[i + 1] = new CustomerItem(headers.get(i).toString(), false);
	        }	
	        
	        combo.setModel(new DefaultComboBoxModel<CustomerItem>(types));
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
	                for (int i = 1; i < types.length; i++) {
	                	if (types[i].status == true) {
	                		types[0].status = false;
	                	}
	                }	
	            }
	        });
	        
	        return combo;
	}
	
	public ArrayList selectedTypes() {
		
		ArrayList selectedValues = new ArrayList<Integer>();
			
    	for (int i = 0; i < types.length; i++) {
    		if (types[i].status == true) {
    			selectedValues.add(types[i].label);
    		}	
    	}
    	
    	return selectedValues;
	}	
}
