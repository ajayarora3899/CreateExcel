package com.samples.excel;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;

public class DaoFile {
	
	private String url="jdbc:mysql://localhost:3306/ComLibrary";
	private String uname="root";
	private String pass="Ajay@1997";
	
	public ResultSet getResult(String query)
	{
		ResultSet rs=null;
		try
		{
			Class.forName("com.mysql.cj.jdbc.Driver");
			Connection con=DriverManager.getConnection(url,uname,pass);
			PreparedStatement st=con.prepareStatement(query);
			rs=st.executeQuery(query);
			return rs;
		}catch(Exception e)
		{
			System.out.println(e);
		}
		return rs;
	}
}
