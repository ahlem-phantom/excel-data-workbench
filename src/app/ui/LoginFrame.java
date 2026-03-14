package app.ui;

import java.awt.BorderLayout;
import java.awt.Color;
import java.awt.Dimension;
import java.awt.FlowLayout;
import java.awt.Font;
import java.awt.GridBagConstraints;
import java.awt.GridBagLayout;
import java.awt.GridLayout;
import java.awt.Insets;
import java.awt.event.ActionEvent;
import java.sql.SQLException;

import app.auth.AuthService;
import app.ui.components.GradientPanel;
import app.ui.components.StyledButton;
import javax.swing.BorderFactory;
import javax.swing.Box;
import javax.swing.BoxLayout;
import javax.swing.JButton;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JPasswordField;
import javax.swing.JTextField;
import javax.swing.SwingConstants;

/**
 * Generic login window with a modern visual layout.
 */
public class LoginFrame extends JFrame {

    private static final long serialVersionUID = 1L;

    private final AuthService auth = new AuthService();

    private JTextField usernameField;
    private JPasswordField passwordField;

    public LoginFrame() {
        setTitle("Excel Insight Studio - Login");
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setSize(940, 560);
        setLocationRelativeTo(null);
        setResizable(false);

        JPanel root = new JPanel(new GridLayout(1, 2));
        root.add(buildHeroPanel());
        root.add(buildFormPanel());
        setContentPane(root);
    }

    private JPanel buildHeroPanel() {
        GradientPanel hero = new GradientPanel(new Color(18, 61, 108), new Color(21, 152, 134));
        hero.setBorder(BorderFactory.createEmptyBorder(38, 44, 38, 44));
        hero.setLayout(new BoxLayout(hero, BoxLayout.Y_AXIS));

        JLabel app = new JLabel("Excel Insight Studio");
        app.setForeground(Color.WHITE);
        app.setFont(new Font("Segoe UI", Font.BOLD, 28));

        JLabel line = new JLabel("Professional Excel Workspace");
        line.setForeground(new Color(221, 244, 255));
        line.setFont(new Font("Segoe UI", Font.PLAIN, 20));

        JLabel point1 = createHeroBullet("Editable data grid with direct cell updates");
        JLabel point2 = createHeroBullet("Search and jump to matching data instantly");
        JLabel point3 = createHeroBullet("Delete rows, undo, and export clean output");

        hero.add(app);
        hero.add(Box.createVerticalStrut(8));
        hero.add(line);
        hero.add(Box.createVerticalStrut(30));
        hero.add(point1);
        hero.add(Box.createVerticalStrut(10));
        hero.add(point2);
        hero.add(Box.createVerticalStrut(10));
        hero.add(point3);
        hero.add(Box.createVerticalGlue());

        return hero;
    }

    private JLabel createHeroBullet(String text) {
        JLabel lbl = new JLabel("\u2713  " + text);
        lbl.setForeground(Color.WHITE);
        lbl.setFont(new Font("Segoe UI", Font.PLAIN, 15));
        return lbl;
    }

    private JPanel buildFormPanel() {
        JPanel shell = new JPanel(new GridBagLayout());
        shell.setBackground(new Color(244, 248, 252));

        JPanel card = new JPanel(new BorderLayout(0, 14));
        card.setPreferredSize(new Dimension(390, 360));
        card.setBackground(Color.WHITE);
        card.setBorder(BorderFactory.createCompoundBorder(
                BorderFactory.createMatteBorder(4, 0, 0, 0, new Color(21, 152, 134)),
                BorderFactory.createEmptyBorder(22, 28, 22, 28)));

        JLabel title = new JLabel("Sign In", SwingConstants.LEFT);
        title.setForeground(new Color(34, 49, 73));
        title.setFont(new Font("Segoe UI", Font.BOLD, 32));
        card.add(title, BorderLayout.NORTH);

        JPanel form = new JPanel(new GridBagLayout());
        form.setOpaque(false);
        GridBagConstraints gbc = new GridBagConstraints();
        gbc.insets = new Insets(8, 0, 8, 0);
        gbc.fill = GridBagConstraints.HORIZONTAL;
        gbc.gridx = 0;
        gbc.gridy = 0;

        form.add(label("Username"), gbc);
        gbc.gridy++;
        usernameField = inputField(20);
        form.add(usernameField, gbc);

        gbc.gridy++;
        form.add(label("Password"), gbc);
        gbc.gridy++;
        passwordField = passwordInput(20);
        form.add(passwordField, gbc);

        card.add(form, BorderLayout.CENTER);

        JPanel actions = new JPanel(new FlowLayout(FlowLayout.RIGHT, 10, 0));
        actions.setOpaque(false);

        JButton registerBtn = secondaryButton("Create Account");
        registerBtn.addActionListener(this::onRegister);

        JButton loginBtn = primaryButton("Login");
        loginBtn.addActionListener(this::onLogin);

        actions.add(registerBtn);
        actions.add(loginBtn);
        card.add(actions, BorderLayout.SOUTH);

        shell.add(card);
        return shell;
    }

    private JLabel label(String text) {
        JLabel lbl = new JLabel(text);
        lbl.setForeground(new Color(56, 69, 89));
        lbl.setFont(new Font("Segoe UI", Font.BOLD, 13));
        return lbl;
    }

    private JTextField inputField(int cols) {
        JTextField tf = new JTextField(cols);
        tf.setFont(new Font("Segoe UI", Font.PLAIN, 14));
        tf.setBorder(BorderFactory.createCompoundBorder(
                BorderFactory.createLineBorder(new Color(205, 215, 228)),
                BorderFactory.createEmptyBorder(8, 10, 8, 10)));
        return tf;
    }

    private JPasswordField passwordInput(int cols) {
        JPasswordField pf = new JPasswordField(cols);
        pf.setFont(new Font("Segoe UI", Font.PLAIN, 14));
        pf.setBorder(BorderFactory.createCompoundBorder(
                BorderFactory.createLineBorder(new Color(205, 215, 228)),
                BorderFactory.createEmptyBorder(8, 10, 8, 10)));
        return pf;
    }

    private JButton primaryButton(String text) {
        StyledButton b = new StyledButton(text, new Color(21, 101, 192), Color.WHITE);
        b.setFont(new Font("Segoe UI", Font.BOLD, 14));
        b.setBorder(BorderFactory.createEmptyBorder(10, 24, 10, 24));
        return b;
    }

    private JButton secondaryButton(String text) {
        StyledButton b = new StyledButton(text, new Color(224, 235, 250), new Color(21, 60, 110));
        b.setFont(new Font("Segoe UI", Font.BOLD, 14));
        b.setBorder(BorderFactory.createEmptyBorder(10, 20, 10, 20));
        return b;
    }

    private void onLogin(ActionEvent e) {
        String username = usernameField.getText().trim();
        String password = new String(passwordField.getPassword());

        if (username.isEmpty()) {
            JOptionPane.showMessageDialog(this, "Please input username.");
            usernameField.requestFocusInWindow();
            return;
        }
        if (password.isEmpty()) {
            JOptionPane.showMessageDialog(this, "Please input password.");
            passwordField.requestFocusInWindow();
            return;
        }

        try {
            boolean ok = auth.authenticate(username, password);
            if (!ok) {
                JOptionPane.showMessageDialog(this, "Incorrect username or password.");
                return;
            }

            MainFrame main = new MainFrame(username);
            main.setVisible(true);
            dispose();

        } catch (SQLException ex) {
            JOptionPane.showMessageDialog(this,
                    "Database error: " + ex.getMessage(),
                    "Login Failed",
                    JOptionPane.ERROR_MESSAGE);
        }
    }

    private void onRegister(ActionEvent e) {
        RegisterFrame registerFrame = new RegisterFrame(this);
        registerFrame.setVisible(true);
        setVisible(false);
    }
}
