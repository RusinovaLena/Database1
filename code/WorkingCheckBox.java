package net.codejava;

import java.awt.Graphics;
import java.awt.Image;
import java.awt.image.BufferedImage;
import java.io.IOException;

import javax.swing.Icon;
import javax.swing.ImageIcon;
import javax.swing.JCheckBox;
import javax.swing.UIManager;

public class WorkingCheckBox {
	// изменение размера галочки в JChecBox
    public void scaleCheckBoxIcon(JCheckBox checkbox, int heightWidth) throws IOException {

        boolean previousState = checkbox.isSelected();
        checkbox.setSelected(false);

        Icon boxIcon = UIManager.getIcon("CheckBox.icon");
        BufferedImage boxImage = new BufferedImage(
                boxIcon.getIconWidth(), boxIcon.getIconHeight(), BufferedImage.TYPE_INT_ARGB);
        Graphics graphics = boxImage.createGraphics();

        try{
            boxIcon.paintIcon(checkbox, graphics, 0, 0);
        } finally
        {
            graphics.dispose();
        }

        ImageIcon newBoxImage = new ImageIcon(boxImage);
        Image finalBoxImage = newBoxImage.getImage().getScaledInstance(
                boxImage.getWidth(), boxImage.getHeight(), Image.SCALE_SMOOTH);
        finalBoxImage = finalBoxImage.getScaledInstance(heightWidth, heightWidth, Image.SCALE_SMOOTH);

        checkbox.setIcon(new ImageIcon(finalBoxImage));
        checkbox.setSelected(true);

        Icon checkedBoxIcon = UIManager.getIcon("CheckBox.icon");
        BufferedImage checkedBoxImage = new BufferedImage(
                boxIcon.getIconWidth(), boxIcon.getIconHeight(), BufferedImage.TYPE_INT_ARGB);
        Graphics checkedGraphics = checkedBoxImage.createGraphics();

        try{
            checkedBoxIcon.paintIcon(checkbox, checkedGraphics, 0, 0);
        } finally{
            checkedGraphics.dispose();
        }

        ImageIcon newCheckedBoxImage = new ImageIcon(checkedBoxImage);
        Image finalCheckedBoxImage = newCheckedBoxImage.getImage().getScaledInstance(boxImage.getWidth(), boxImage.getHeight(), Image.SCALE_SMOOTH);
        finalCheckedBoxImage = finalCheckedBoxImage.getScaledInstance(heightWidth, heightWidth, Image.SCALE_SMOOTH);

        checkbox.setSelectedIcon(new ImageIcon(finalCheckedBoxImage));
        checkbox.setSelected(false);
        checkbox.setSelected(previousState);
    }
}
