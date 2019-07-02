package port;

import com.codoid.products.exception.FilloException;
import com.codoid.products.fillo.Connection;
import com.codoid.products.fillo.Fillo;
import com.codoid.products.fillo.Recordset;

public class ReadExcel {
	public static void main(String args[]) throws FilloException{
		Fillo fillo=new Fillo();
		Connection conn=null;
		try {
			conn=fillo.getConnection("C:\\Users\\User1\\Downloads\\Book 3.xlsx");
			
		}catch(FilloException e) {
			throw new FilloException("File not found");
			
			}
		String query = "Select * from sheet1 ";
		Recordset rcrdset=null;
		try {
			rcrdset=conn.executeQuery(query);
		
		}catch(FilloException e) {
			throw new FilloException("Error executing query");
			
		}
		try {
			System.out.println("Username"+ "\t" + "Password" + "\t\t" + "Mail_id"+  "\t\t\t\t" +"Mobile_no");
			while(rcrdset.next()) {
				
				System.out.print(rcrdset.getField("User_name")+"\t\t ");
				System.out.print(rcrdset.getField("Password")+"\t\t");
				System.out.print(rcrdset.getField("Mail_id")+" \t\t");
				System.out.print(rcrdset.getField("Mobile_no")+"\t\t");
				System.out.println("\n");
			}
			
			
		
		}catch(FilloException e) {
		throw new FilloException("No records found");
		
		}finally {
			rcrdset.close();
			conn.close();
		}
		
	
				
			
		
	}
	

}
