package app.auth;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;

import org.mindrot.jbcrypt.BCrypt;

/**
 * Handles user authentication against a MySQL database.
 *
 * All queries use PreparedStatements to prevent SQL injection.
 * Passwords are stored as BCrypt hashes. 
 */
public class AuthService {

    private static final String DB_ROOT_URL = "jdbc:mysql://localhost:3306/";
    private static final String DB_URL = "jdbc:mysql://localhost:3306/excel_viewer_db";
    private static final String DB_USER = "root";
    private static final String DB_PASS = "";

    private static final int BCRYPT_COST = 12;

    private static boolean schemaReady;

    static {
        try {
            Class.forName("com.mysql.cj.jdbc.Driver");
        } catch (ClassNotFoundException e) {
            throw new RuntimeException("MySQL JDBC driver not found on classpath.", e);
        }
    }

    /**
     * Registers a new user. Throws SQLException if the username already exists.
     */
    public void register(String username, String email, String password) throws SQLException {
        ensureSchemaReady();
        final String sql = "INSERT INTO users (username, email, password) VALUES (?, ?, ?)";
        try (Connection con = connect();
             PreparedStatement ps = con.prepareStatement(sql)) {
            ps.setString(1, username);
            ps.setString(2, email);
            ps.setString(3, hashPassword(password));
            ps.executeUpdate();
        }
    }

    /**
     * Returns true if the supplied credentials match a record in the database.
     */
    public boolean authenticate(String username, String password) throws SQLException {
        ensureSchemaReady();
        final String sql = "SELECT id, password FROM users WHERE username = ?";
        try (Connection con = connect();
             PreparedStatement ps = con.prepareStatement(sql)) {
            ps.setString(1, username);
            try (ResultSet rs = ps.executeQuery()) {
                if (!rs.next()) {
                    return false;
                }

                String storedHash = rs.getString("password");
                if (storedHash == null || storedHash.trim().isEmpty()) {
                    return false;
                }

                try {
                    return BCrypt.checkpw(password, storedHash);
                } catch (IllegalArgumentException invalidHash) {
                    return false;
                }
            }
        }
    }

    private Connection connect() throws SQLException {
        return DriverManager.getConnection(DB_URL, DB_USER, DB_PASS);
    }

    private static synchronized void ensureSchemaReady() throws SQLException {
        if (schemaReady) {
            return;
        }
        ensureSchema();
        schemaReady = true;
    }

    private static void ensureSchema() throws SQLException {
        try (Connection con = DriverManager.getConnection(DB_ROOT_URL, DB_USER, DB_PASS);
             Statement st = con.createStatement()) {
            st.executeUpdate("CREATE DATABASE IF NOT EXISTS excel_viewer_db");
        } catch (SQLException ignored) {
            // If the database already exists and this user only has per-database
            // privileges, the connection below is enough.
        }

        try (Connection con = DriverManager.getConnection(DB_URL, DB_USER, DB_PASS);
             Statement st = con.createStatement()) {
            st.executeUpdate(
                    "CREATE TABLE IF NOT EXISTS users (" +
                    "id INT AUTO_INCREMENT PRIMARY KEY, " +
                    "username VARCHAR(100) NOT NULL UNIQUE, " +
                    "email VARCHAR(150) NOT NULL, " +
                    "password VARCHAR(60) NOT NULL)");
            st.executeUpdate("ALTER TABLE users MODIFY password VARCHAR(60) NOT NULL");
        }
    }

    private static String hashPassword(String password) {
        return BCrypt.hashpw(password, BCrypt.gensalt(BCRYPT_COST));
    }
}
