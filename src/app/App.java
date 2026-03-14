package app;

import java.awt.EventQueue;
import javax.swing.UIManager;

import app.ui.LoginFrame;

/**
 * Generic application entry point.
 */
public class App {

    public static void main(String[] args) {
        EventQueue.invokeLater(() -> {
            try {
                UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
            } catch (Exception ignored) {
                // Fallback to default look and feel
            }
            new LoginFrame().setVisible(true);
        });
    }
}
