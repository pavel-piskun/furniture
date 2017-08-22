package by.mebel.db.h2;
import java.sql.*;

public class testH2 {
    public static void main(String[] a)
            throws Exception {
        Class.forName("org.h2.Driver");
        System.out.println("get connection");
        Connection conn = DriverManager.
                getConnection("jdbc:h2:data/db/test", "sa", "");

        Statement statement = conn.createStatement();
        /*statement.execute("CREATE TABLE USER (ID INT, NAME VARCHAR(50));");
        conn.commit();
        statement.execute("INSERT INTO USER VALUES(1, 'PISKUN'); ");
        conn.commit();
        */
        ResultSet rs = statement.executeQuery("SELECT * FROM USER");
        
        while (rs.next()) {
            System.out.println("ID: "+rs.getInt("ID")+" NAME: "+rs.getString("NAME"));            
        }
        conn.close();
        System.out.println("end");
    }
}
