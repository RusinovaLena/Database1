package net.codejava;

import java.util.Calendar;

import javax.swing.JCheckBox;

public class DateAnalysis {
	boolean nRowRed = false;
	boolean nRowYellow = false;
	String sD;
	
	DateAnalysis() {}
	
	public DateAnalysis (boolean nRowRed, boolean nRowYellow, String sD) {
		this.nRowRed = nRowRed;
		this.nRowYellow = nRowYellow;
		this.sD = sD;
	}
	String getSD() {
		return sD;
	}
	public boolean checkString(Calendar c1, Calendar c2, JCheckBox field1) {
		boolean b2 = true;
		boolean sameDay = false;
		if (c2.get(Calendar.YEAR) < c2.get(Calendar.YEAR)) {                                          	
            sameDay = true;
        } else {

            if (c2.get(Calendar.YEAR) > c2.get(Calendar.YEAR)) {
                sameDay = false;
            } else {

                if (c2.get(Calendar.MONTH) < c2.get(Calendar.MONTH)) {
                	 sameDay = true;                   
                    if (c2.get(Calendar.MONTH) - c2.get(Calendar.MONTH) == 1) {                                                           	
                    	if (30 - (c2.get(Calendar.DAY_OF_MONTH)) +  c2.get(Calendar.DAY_OF_MONTH) <= 30) {
                    		nRowYellow = true;
                    	} else {
                    		nRowYellow = false;
                    	}
                        	
                    	if (30 - (c2.get(Calendar.DAY_OF_MONTH)) +  c2.get(Calendar.DAY_OF_MONTH) <= 7) {
                    		nRowRed = true;
                    	} else {
                    		nRowRed = false;
                    	}
                	}
                } else {                                                       	
                    if (c2.get(Calendar.MONTH) > c2.get(Calendar.MONTH)) {
                        sameDay = false;
                                                                                   
                    } else {                                                       	            	                                                            	                                                                        	
                        if (c2.get(Calendar.DAY_OF_MONTH) <= c2.get(Calendar.DAY_OF_MONTH)) {
                            sameDay = true;
                        	nRowYellow = true;
                        	if (c2.get(Calendar.DAY_OF_MONTH) - c2.get(Calendar.DAY_OF_MONTH) <= 7) {
                        		nRowRed = true;
                        	} else {
                        		nRowRed = false;
                        	}
                        } else {
                            sameDay = false;   	                                                                             
                        }
                    }
                }
            }
        }    
        sD = String.valueOf(sameDay);
        
        if (sD == String.valueOf(field1.isSelected())) {
        	b2 = true;
        } else {
        	b2 = false;
        }
		return b2;
	}
	
	public boolean checkNum(Calendar c1, Calendar c2, JCheckBox field1) {
		boolean b2 = false;
		boolean sameDay = false;
		if (c1.get(Calendar.YEAR) < c2.get(Calendar.YEAR)) {
        	sameDay = true; 
        } else {
            if (c1.get(Calendar.YEAR) > c2.get(Calendar.YEAR)) {
                sameDay = false;
            } else {
                if (c1.get(Calendar.MONTH) < c2.get(Calendar.MONTH)) {
                	sameDay = true;
                	
                    if (c2.get(Calendar.MONTH) - c1.get(Calendar.MONTH) == 1) {                                                     	
                    	if ((30 - c1.get(Calendar.DAY_OF_MONTH)) +  c2.get(Calendar.DAY_OF_MONTH) <= 30) {
                    		nRowYellow = true;
                    	} else {
                    		nRowYellow = false;
                    	}
                    	
                    	if (30 - (c1.get(Calendar.DAY_OF_MONTH)) +  c2.get(Calendar.DAY_OF_MONTH) <= 7) {
                    		nRowRed = true;
                    	} else {
                    		nRowRed = false;
                    	}
                	}                                                                     
                } else {

                    if (c1.get(Calendar.MONTH) > c2.get(Calendar.MONTH)) {
                        sameDay = false;
                        
                    } else {                                                       	
                        if (c1.get(Calendar.DAY_OF_MONTH) <= c2.get(Calendar.DAY_OF_MONTH)) {
                            sameDay = true;
                            nRowYellow = true;
                            
                        	if (c2.get(Calendar.DAY_OF_MONTH) - c1.get(Calendar.DAY_OF_MONTH) <= 7) {
                        		nRowRed = true;
                        	} else {
                        		nRowRed = false;
                        	}
                        } else {
                            sameDay = false;  	                                                                                        
                        }
                    }
                }
            }
        }                                                                       
        return b2;
	}   
	
}
