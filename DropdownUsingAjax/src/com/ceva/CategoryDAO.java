package com.ceva;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.List;

public class CategoryDAO {
	String driver="org.postgresql.Driver";
    String databaseURL = "jdbc:postgresql://localhost:5432/postgres";
    String user = "postgres";
    String password = "root";
     
    public List<Category> list() throws SQLException {
        List<Category> listCategory = new ArrayList<>();
         
        try{
        	Class.forName(driver);
        	Connection connection = DriverManager.getConnection(databaseURL, user, password);
        
            String sql = "SELECT * FROM category ORDER BY name";
            Statement statement = connection.createStatement();
            ResultSet result = statement.executeQuery(sql);
             
            while (result.next()) {
                int id = result.getInt("category_id");
                String name = result.getString("name");
                Category category = new Category(id, name);
                     
                listCategory.add(category);
            }          
             
        } catch (SQLException | ClassNotFoundException ex) {
            ex.printStackTrace();
        }      
         
        return listCategory;
    }
}
