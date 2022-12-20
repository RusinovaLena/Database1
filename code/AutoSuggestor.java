package net.codejava;

import java.awt.BorderLayout;
import java.awt.Color;
import java.awt.Dimension;
import java.awt.GridLayout;
import java.awt.Window;
import java.awt.event.ActionEvent;
import java.awt.event.KeyEvent;
import java.util.ArrayList;

import javax.swing.AbstractAction;
import javax.swing.JComponent;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.swing.JScrollPane;
import javax.swing.JTextField;
import javax.swing.JWindow;
import javax.swing.KeyStroke;
import javax.swing.event.DocumentEvent;
import javax.swing.event.DocumentListener;

// ����� ��� ������ ���������� ����� ���������� ���������� � ��
public class AutoSuggestor {
	public final JTextField textField;
    private final Window container;
    private JPanel suggestionsPanel;
    public JWindow autoSuggestionPopUpWindow ;
    public String typedWord;
    public ArrayList<String> dictionary = new ArrayList<>();
    private int tW, tH;
    static int cH = 0;
    static int cW = 0;
    public DocumentListener documentListener = new DocumentListener() {
    	
        @Override
        public void insertUpdate(DocumentEvent de) {
            checkForAndShowSuggestions();       
        }

        @Override
        public void removeUpdate(DocumentEvent de) {
            checkForAndShowSuggestions();
        }

        @Override
        public void changedUpdate(DocumentEvent de) {
            checkForAndShowSuggestions();
        }
        
    };
    
    private final Color suggestionsTextColor;
    private final Color suggestionFocusedColor;

    public AutoSuggestor(JTextField textField, Window mainWindow, ArrayList<String> words, 
    		Color popUpBackground, Color textColor, Color suggestionFocusedColor, float opacity, int cH, int cW) {   	
    	checkForAndShowSuggestions();   	
    	this.cH = cH;
    	this.cW = cW;
    	this.textField = textField;
        this.suggestionsTextColor = textColor;
        this.container = mainWindow;
        this.suggestionFocusedColor = suggestionFocusedColor;
        this.textField.getDocument().addDocumentListener(documentListener);      
        typedWord = "";
        tW = 0;
        tH = 0;
        autoSuggestionPopUpWindow = new JWindow(mainWindow);
        autoSuggestionPopUpWindow.setOpacity(opacity);
        suggestionsPanel = new JPanel();
        suggestionsPanel.setLayout(new GridLayout(0, 1));
        suggestionsPanel.setBackground(popUpBackground);
    }

    private void setFocusToTextField() {
        container.toFront();
        container.requestFocusInWindow();
        textField.requestFocusInWindow();
    }

    public ArrayList<SuggestionLabel> getAddedSuggestionLabels() {
        ArrayList<SuggestionLabel> sls = new ArrayList<>();
        for (int i = 0; i < suggestionsPanel.getComponentCount(); i++) {
            if (suggestionsPanel.getComponent(i) instanceof SuggestionLabel) {
                SuggestionLabel sl = (SuggestionLabel) suggestionsPanel.getComponent(i);
                sls.add(sl);
            }
        }
        return sls;
    }

    void checkForAndShowSuggestions() {
    	
        typedWord = getCurrentlyTypedWord();   
        if (suggestionsPanel != null) {
        	suggestionsPanel.removeAll();
        }	
  
        tW = 0;
        tH = 0;
        boolean added = wordTyped(typedWord);        
        if (!added) {      	
        	if (autoSuggestionPopUpWindow == null) {
        	} else {
	            if (autoSuggestionPopUpWindow.isVisible()) {
	                autoSuggestionPopUpWindow.setVisible(true);	     
	            }
        	}    
        } else {        	
            showPopUpWindow();
            setFocusToTextField();
        }
    }
    
    protected void addWordToSuggestions(String word) {
    	
        SuggestionLabel suggestionLabel = new SuggestionLabel(word, suggestionFocusedColor, suggestionsTextColor, this);
        calculatePopUpWindowSize(suggestionLabel);
        suggestionsPanel.add(suggestionLabel);               
    }
    
    public String getCurrentlyTypedWord() {
    	String text;   	
    	if (textField != null) {
    		text = textField.getText();
    	} else {
    		text = "";
    	}

        return text;

    }

    private void calculatePopUpWindowSize(JLabel label) {

        if (tW < label.getPreferredSize().width) {
            tW = label.getPreferredSize().width;
        }
        tH += label.getPreferredSize().height;
    }

    private void showPopUpWindow() {
    	
    	
    	
        autoSuggestionPopUpWindow.getContentPane().add(suggestionsPanel);
        autoSuggestionPopUpWindow.setMinimumSize(new Dimension(textField.getWidth(), 40));
        autoSuggestionPopUpWindow.setSize(tW, tH);
        autoSuggestionPopUpWindow.setVisible(true);
        
        int windowX = 0;
        int windowY = 0;
        windowX = container.getX() + textField.getX();
        
        if (suggestionsPanel.getHeight() > autoSuggestionPopUpWindow.getMinimumSize().height) {
            windowY = container.getY() + textField.getY() + textField.getHeight() + autoSuggestionPopUpWindow.getMinimumSize().height;
        } else {
            windowY = container.getY() + textField.getY() + textField.getHeight() + autoSuggestionPopUpWindow.getHeight();
        }
        System.out.println(windowX + " w " + windowY);
        System.out.println(container.getY() + " g " + container.getX());
        System.out.println(textField.getY() + " t " + textField.getX());
        System.out.println(tW + " hw " + tH);
        System.out.println(autoSuggestionPopUpWindow.getLocale() + " hw " + autoSuggestionPopUpWindow.getLocation() );
        
        autoSuggestionPopUpWindow.setLocation(windowX + cW, windowY + cH);
        int a1 = windowX + 665;
        int a2 = windowY + 168;
        System.out.println( (windowX + 810) + " r " + (windowY + 600) );
        
        autoSuggestionPopUpWindow.setMinimumSize(new Dimension(textField.getWidth(), 40));
        autoSuggestionPopUpWindow.revalidate();
        autoSuggestionPopUpWindow.repaint();
    }

    public void setDictionary(ArrayList<String> words) {
    	
        dictionary.clear();
        if (words == null) {
        } else {
	        for (String word : words) {
	            dictionary.add(word);
	        }
        }   
    }


    public JWindow getAutoSuggestionPopUpWindow() {
        return autoSuggestionPopUpWindow;
       
    }

    public Window getContainer() {
    	
        return container;
    }

    public JTextField getTextField() {
        return textField;
    }

    public void addToDictionary(String word) {
        dictionary.add(word);
    }
    
    boolean wordTyped(String typedWord) {

        if (typedWord.isEmpty()) {
    		
    		if (autoSuggestionPopUpWindow != null) {
    			
    			if (autoSuggestionPopUpWindow.isVisible()) {
                    autoSuggestionPopUpWindow.setVisible(false);	              
                }
    		}	   		
            return false;
        }
        
        boolean suggestionAdded = false;
        
        boolean fullymatches = true;
        for (String word : dictionary) {
         
            if (typedWord.equals(word)) {  
            	autoSuggestionPopUpWindow.setVisible(false);
            	fullymatches = false;
            } else {
	            for (int i = 0; i < typedWord.length(); i++) { 	 
	            	if (typedWord.length() < word.length()) {
		                if (!typedWord.toLowerCase().
		                	startsWith(String.valueOf(word.toLowerCase().charAt(i)), i)) {
		                	fullymatches = false;
		                	break;
		                }
	            	} else {
	            		 fullymatches = false;
	            		 break;
	            	}
	            }	
            }
            if (fullymatches) {            		
        		addWordToSuggestions(word);
        		suggestionAdded = true;         	
            } else {
            	fullymatches = true;
            }
        }
        return suggestionAdded;
    }
}

