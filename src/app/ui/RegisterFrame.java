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
 * Generic user registration form with a richer visual style.
 */
public class RegisterFrame extends JFrame {

    private static final long serialVersionUID = 1L;

    private final AuthService auth = new AuthService();
    private final LoginFrame loginFrame;

    private JTextField usernameField;
    private JTextField emailField;
    private JPasswordField passwordField;
    private JPasswordField confirmPasswordField;

    public RegisterFrame(LoginFrame loginFrame) {
        this.loginFrame = loginFrame;

        setTitle("Excel Insight Studio - Register");
        setDefaultCloseOperation(JFrame.DISPOSE_ON_CLOSE);
        setSize(980, 620);
        setLocationRelativeTo(null);
        setResizable(false);

        addWindowListener(new java.awt.event.WindowAdapter() {
            @Override
            public void windowClosing(java.awt.event.WindowEvent e) {
                loginFrame.setVisible(true);
            }
        });

        JPanel root = new JPanel(new GridLayout(1, 2));
        root.add(buildHeroPanel());
        root.add(buildFormPanel());
        setContentPane(root);
    }

    private JPanel buildHeroPanel() {
        GradientPanel hero = new GradientPanel(new Color(17, 76, 123), new Color(33, 170, 136));
        hero.setLayout(new BoxLayout(hero, BoxLayout.Y_AXIS));
        hero.setBorder(BorderFactory.createEmptyBorder(44, 44, 44, 44));

        JLabel title = new JLabel("Create Account");
        title.setForeground(Color.WHITE);
        title.setFont(new Font("Segoe UI", Font.BOLD, 28));

        JLabel subtitle = new JLabel("Get started with secure Excel workflows.");
        subtitle.setForeground(new Color(227, 246, 255));
        subtitle.setFont(new Font("Segoe UI", Font.PLAIN, 20));

        JLabel r1 = heroLine("BCrypt-secured user accounts");
        JLabel r2 = heroLine("Smart .xlsx import with header detection");
        JLabel r3 = heroLine("Edit, search, undo, and export fast");

        hero.add(title);
        hero.add(Box.createVerticalStrut(8));
        hero.add(subtitle);
        hero.add(Box.createVerticalStrut(28));
        hero.add(r1);
        hero.add(Box.createVerticalStrut(12));
        hero.add(r2);
        hero.add(Box.createVerticalStrut(12));
        hero.add(r3);
        hero.add(Box.createVerticalGlue());
        return hero;
    }

    private JLabel heroLine(String text) {
        JLabel lbl = new JLabel("\u2713  " + text);
        lbl.setForeground(Color.WHITE);
        lbl.setFont(new Font("Segoe UI", Font.PLAIN, 15));
        return lbl;
    }

    private JPanel buildFormPanel() {
        JPanel shell = new JPanel(new GridBagLayout());
        shell.setBackground(new Color(243, 247, 251));

        JPanel card = new JPanel(new BorderLayout(0, 14));
        card.setPreferredSize(new Dimension(420, 500));
        card.setBackground(Color.WHITE);
        card.setBorder(BorderFactory.createCompoundBorder(
                BorderFactory.createMatteBorder(4, 0, 0, 0, new Color(21, 152, 134)),
                BorderFactory.createEmptyBorder(22, 28, 22, 28)));

        JLabel heading = new JLabel("Register", SwingConstants.LEFT);
        heading.setFont(new Font("Segoe UI", Font.BOLD, 32));
        heading.setForeground(new Color(34, 49, 73));
        card.add(heading, BorderLayout.NORTH);

        JPanel form = new JPanel(new GridBagLayout());
        form.setOpaque(false);
        GridBagConstraints gbc = new GridBagConstraints();
        gbc.insets = new Insets(7, 0, 7, 0);
        gbc.fill = GridBagConstraints.HORIZONTAL;
        gbc.gridx = 0;
        gbc.gridy = 0;

        form.add(label("Username"), gbc);
        gbc.gridy++;
        usernameField = inputField(22);
        form.add(usernameField, gbc);

        gbc.gridy++;
        form.add(label("Email"), gbc);
        gbc.gridy++;
        emailField = inputField(22);
        form.add(emailField, gbc);

        gbc.gridy++;
        form.add(label("Password"), gbc);
        gbc.gridy++;
        passwordField = passwordInput(22);
        form.add(passwordField, gbc);

        gbc.gridy++;
        form.add(label("Confirm Password"), gbc);
        gbc.gridy++;
        confirmPasswordField = passwordInput(22);
        form.add(confirmPasswordField, gbc);

        card.add(form, BorderLayout.CENTER);

        JPanel actions = new JPanel(new FlowLayout(FlowLayout.RIGHT, 10, 0));
        actions.setOpaque(false);

        JButton backBtn = secondaryButton("Back");
        backBtn.addActionListener(this::onBack);

        JButton registerBtn = primaryButton("Register");
        registerBtn.addActionListener(this::onRegister);

        actions.add(backBtn);
        actions.add(registerBtn);
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

    private void onBack(ActionEvent e) {
        loginFrame.setVisible(true);
        dispose();
    }

    private void onRegister(ActionEvent e) {
        String username = usernameField.getText().trim();
        String email = emailField.getText().trim();
        String password = new String(passwordField.getPassword());
        String confirm = new String(confirmPasswordField.getPassword());

        if (username.length() < 4) {
            JOptionPane.showMessageDialog(this, "Username must be at least 4 characters.");
            usernameField.requestFocusInWindow();
            return;
        }
        if (!isValidEmail(email)) {
            JOptionPane.showMessageDialog(this, "Please enter a valid email address.");
            emailField.requestFocusInWindow();
            return;
        }
        if (password.length() < 6) {
            JOptionPane.showMessageDialog(this, "Password must be at least 6 characters.");
            passwordField.requestFocusInWindow();
            return;
        }
        if (!password.equals(confirm)) {
            JOptionPane.showMessageDialog(this, "Passwords do not match.");
            confirmPasswordField.requestFocusInWindow();
            return;
        }

        try {
            auth.register(username, email, password);
            JOptionPane.showMessageDialog(this, "Registration successful. Please login.");
            loginFrame.setVisible(true);
            dispose();
        } catch (SQLException ex) {
            JOptionPane.showMessageDialog(this,
                    "Could not register user: " + ex.getMessage(),
                    "Registration Failed",
                    JOptionPane.ERROR_MESSAGE);
        }
    }

    private boolean isValidEmail(String email) {
        return email != null && email.matches("^[A-Za-z0-9._%+\\-]+@[A-Za-z0-9.\\-]+\\.[A-Za-z]{2,}$");
    }
}
