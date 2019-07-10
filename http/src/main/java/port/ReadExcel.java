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
			conn=fillo.getConnection("C:\\Users\\User1\\Downloads\\data1.xlsx");
			
		}catch(FilloException e) {
			throw new FilloException("File not found");
			
			}
		String query = "Select * from data1 ";
		Recordset rcrdset=null;
		try {
			rcrdset=conn.executeQuery(query);
		
		}catch(FilloException e) {
			throw new FilloException("Error executing query");
			
		}
		try {
			System.out.println("Date" + "\t\t\t" + "Open" + "\t\t\t" + "High"+  "\t\t\t" +"Low"+"\t\t"+"Close"+"\t\t"+"Shares Traded");
			while(rcrdset.next()) {
				
				System.out.print(rcrdset.getField("Date")+"\t\t ");
				System.out.print(rcrdset.getField("Open")+"\t\t");
				System.out.print(rcrdset.getField("High")+" \t\t");
				System.out.print(rcrdset.getField("Low")+"\t\t");
				System.out.print(rcrdset.getField("close")+"\t\t");
				System.out.print(rcrdset.getField("shares Traded")+"\t\t");
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