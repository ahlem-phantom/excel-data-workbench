package app.ui.components;

import java.awt.Color;
import java.awt.Cursor;
import java.awt.Font;
import java.awt.Graphics;
import java.awt.Graphics2D;
import java.awt.RenderingHints;
import javax.swing.BorderFactory;
import javax.swing.JButton;

/**
 * A JButton that reliably paints a solid rounded-rectangle background regardless
 * of the active Swing Look-and-Feel.
 */
public class StyledButton extends JButton {

    private static final long serialVersionUID = 1L;

    public StyledButton(String text, Color background, Color foreground) {
        super(text);
        setBackground(background);
        setForeground(foreground);
        setFont(new Font("Segoe UI", Font.BOLD, 13));
        setFocusPainted(false);
        setOpaque(false);
        setContentAreaFilled(false);
        setBorderPainted(false);
        setBorder(BorderFactory.createEmptyBorder(9, 18, 9, 18));
        setCursor(Cursor.getPredefinedCursor(Cursor.HAND_CURSOR));
    }

    @Override
    protected void paintComponent(Graphics g) {
        Graphics2D g2 = (Graphics2D) g.create();
        g2.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);

        Color bg;
        if (!isEnabled()) {
            bg = new Color(180, 185, 195);
        } else if (getModel().isPressed()) {
            bg = getBackground().darker();
        } else if (getModel().isRollover()) {
            bg = brighten(getBackground());
        } else {
            bg = getBackground();
        }

        g2.setColor(bg);
        g2.fillRoundRect(0, 0, getWidth(), getHeight(), 12, 12);
        g2.dispose();

        // Let Swing paint the label (text + icon) on top
        super.paintComponent(g);
    }

    @Override
    protected void paintBorder(Graphics g) {
        // Intentionally empty – the rounded fill is sufficient
    }

    private Color brighten(Color c) {
        float[] hsb = Color.RGBtoHSB(c.getRed(), c.getGreen(), c.getBlue(), null);
        return Color.getHSBColor(hsb[0], Math.max(0f, hsb[1] - 0.06f), Math.min(1f, hsb[2] + 0.13f));
    }
}
