
import java.io.File;
import java.io.IOException;
import java.sql.*;

import jxl.read.biff.BiffException;
import jxl.write.WriteException;

public class JExcelAPIDemo
{
   public static void main(String[] args) 
      throws BiffException, IOException, WriteException, IllegalAccessException, InstantiationException, ClassNotFoundException, SQLException
   {
	  String dateStr = "(Sep-07-2015)";
	  
	   
	  StaticMethods.morningstarPopulate(new File("MorningstarStocks" + dateStr + ".xls"));
	  StaticMethods.schwabPopulate(new File("SchwabStocks" + dateStr + ".xls"));
	  StaticMethods.barronsPopulate(new File ("BarronsStocks" + dateStr + ".xls"));
	  StaticMethods.generateResults(new File("StockResults" + dateStr + ".xls"));
	   
   }
}