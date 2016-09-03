import java.io.File;
import java.io.IOException;
import java.sql.*;
import java.util.Calendar;
import java.util.Date;

import jxl.Cell;
import jxl.CellView;
import jxl.Sheet;
import jxl.Workbook;
import jxl.format.CellFormat;
import jxl.read.biff.BiffException;
import jxl.write.Blank;
import jxl.format.Colour;
import jxl.write.Label;
import jxl.write.Number;
import jxl.write.NumberFormats;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;


public final class StaticMethods {

	private static int testInt;
	private static final String url = "jdbc:mysql://localhost:3306/";
	private static final String dbName = "stocks";
    private static final String driver = "com.mysql.jdbc.Driver";
    private static final String userName = "root";
    private static final Date lastBusinessDate = new Date(114, 11, 14, 17, 0);
	
	private StaticMethods(){
		testInt = 7;
	}

	public static void morningstarPopulate(File fille)		//Populates database with Morningstar stocks and ratings
		throws BiffException, IOException, WriteException, IllegalAccessException, InstantiationException, ClassNotFoundException, SQLException{		
		
		Workbook workbook = Workbook.getWorkbook(fille);			//Morningstar Stocks Excel file
	    Sheet sheet = workbook.getSheet(0);
	    Class.forName(driver).newInstance();
	    Connection conn = DriverManager.getConnection(url+dbName, userName, "");		//Database connection
	    for (int i = 1; i < sheet.getRows()-1; i++){
	    	try{
		    	//A stock object is created from reading Morningstar Excel sheet's data
		    	Stock stock = new Stock(sheet.getCell(0, i).getContents(), sheet.getCell(1, i).getContents());
		    	stock.setAnnualYield(Double.parseDouble(sheet.getCell(2, i).getContents()));
		    	stock.setClosePrice(Double.parseDouble(sheet.getCell(3, i).getContents()));
		    	//DateAdded generates automatically
		    	stock.setStarRating(Short.parseShort(sheet.getCell(5, i).getContents()));
		    	stock.setFairValue(Double.parseDouble(sheet.getCell(7, i).getContents()));
		    	stock.setConsiderBuyingPrice(Double.parseDouble(sheet.getCell(8, i).getContents()));
		    	stock.setSector(sheet.getCell(9, i).getContents());
		    	stock.setIndustry(sheet.getCell(10, i).getContents());
		    	stock.setMarketCap(Double.parseDouble(sheet.getCell(11, i).getContents()));
		    	stock.setBusinessDate(lastBusinessDate);
		    	System.out.println(stock.toString());
		    	
		    	Statement st = conn.createStatement(); 
				ResultSet res = st.executeQuery("SELECT * FROM  Stock WHERE Ticker= '" + stock.getTicker() + "';");
				if (res.next()){				
					//Update stocks already in Stock table
					String update = "UPDATE Stock SET ";
					update += "Description='" + stock.getDescription() + "', ";
					update += "AnnualYield=" + stock.getAnnualYield() + ", ";
					update += "ClosePrice=" + stock.getClosePrice() + ", ";
					update += "Sector='" + stock.getSector() + "', ";
					update += "Industry='" + stock.getIndustry() + "', ";
					update += "MarketCap=" + stock.getMarketCap() + " ";
					update += " WHERE Ticker='" + stock.getTicker() + "';";		
					int passTest = st.executeUpdate(update);
					if (passTest==1){System.out.println("Stock " + stock.getTicker() + " updated.");}
					else{System.out.println("UPDATE MORNINGSTAR STOCKS FAILLLL!");}
				}
				else {
					//Insert new stocks into Stock table
					String insert = "INSERT INTO Stock ";
					insert += "(Ticker, Description, AnnualYield, ClosePrice, DateAdded, Sector, Industry, MarketCap) ";
					insert += "VALUES (";
					insert += "'" + stock.getTicker() + "', ";
					insert += "'" + stock.getDescription() + "', ";
					insert += stock.getAnnualYield() + ", ";
					insert += stock.getClosePrice() + ", ";
					insert += "'" + convertJavaDateToSqlDate(stock.getDateAdded()).toString() + "', ";
					insert += "'" + stock.getSector() + "', ";
					insert += "'" + stock.getIndustry() + "', ";
					insert += stock.getMarketCap() + ");";
					int passTest = st.executeUpdate(insert);
					if (passTest==1){System.out.println("Stock " + stock.getTicker() + " added.");}
					else{System.out.println("INSERT MORNINGSTAR STOCKS FAILLLL!");}
				}
				
				ResultSet msRes = st.executeQuery("SELECT * FROM  Morningstar_Score WHERE Ticker= '" + stock.getTicker() + "';");
				if (msRes.next()){
					//Update stocks already in Morningstar_Score table
					String update = "UPDATE Morningstar_Score SET ";
					update += "StarRating=" + stock.getStarRating() + ", ";
					update += "FairValue=" + stock.getFairValue() + ", ";
					update += "ConsiderBuyPrice=" + stock.getConsiderBuyingPrice() + ", ";
					update += "BusinessDate='" + convertJavaDateToSqlDate(stock.getBusinessDate()).toString() + "' "; 
					update += " WHERE Ticker='" + stock.getTicker() + "';";		
					int passTest = st.executeUpdate(update);
					if (passTest==1){System.out.println("Morningstar score for " + stock.getTicker() + " updated.");}
					else{System.out.println("UPDATE MORNINGSTAR_SCORE STOCKS FAILLLL!");}
				}
				else {
					//Insert new stocks into Morningstar_Score table
					String insert = "INSERT INTO Morningstar_Score (Ticker, StarRating, FairValue, ConsiderBuyPrice, BusinessDate) VALUES (";
					insert += "'" + stock.getTicker() + "', ";
					insert += stock.getStarRating() + ", ";
					insert += stock.getFairValue() + ", ";
					insert += stock.getConsiderBuyingPrice()+ ", ";
					insert += "'" + convertJavaDateToSqlDate(stock.getBusinessDate()).toString() + "');";
					int passTest = st.executeUpdate(insert);
					if (passTest==1){System.out.println("Morningstar score for " + stock.getTicker() + " added.");}
					else{System.out.println("INSERT MORNINGSTAR_SCORE STOCKS FAILLLL!");}
				}
	    	}
	    	catch (Exception e){
	    		e.printStackTrace();
	    		conn.close();
	    		workbook.close();
	    	}
	    	
	    	
	    	
	    }
	    conn.close();
	    workbook.close();
		
	}
	
	public static void schwabPopulate(File fille)			//Populates database with Schwab stocks and ratings
			throws BiffException, IOException, WriteException, IllegalAccessException, InstantiationException, ClassNotFoundException, SQLException{		
			
			Workbook workbook = Workbook.getWorkbook(fille);			//Schwab Stocks Excel file
		    Sheet sheet = workbook.getSheet(0);
		    Class.forName(driver).newInstance();
		    Connection conn = DriverManager.getConnection(url+dbName, userName, "");		//Database connection
		    for (int i = 1; i < sheet.getRows(); i++){
		    	try{
			    	//A stock object is created from reading Schwab Excel sheet's data
			    	Stock stock = new Stock(sheet.getCell(0, i).getContents(), sheet.getCell(1, i).getContents());
			    	stock.setAnnualYield(round(Double.parseDouble(sheet.getCell(2, i).getContents()) * 100, 2));
			    	stock.setSpRating(Short.parseShort(sheet.getCell(3, i).getContents()));			    	
			    	stock.setClosePrice(Double.parseDouble(sheet.getCell(4, i).getContents()));
			    	//DateAdded generates automatically
			    	stock.setBusinessDate(lastBusinessDate);
			    	//System.out.println(stock.toString());
			    	
			    	Statement st = conn.createStatement(); 
					ResultSet res = st.executeQuery("SELECT * FROM  Stock WHERE Ticker= '" + stock.getTicker() + "';");
					if (res.next()){				
						//Update stocks already in Stock table
						String update = "UPDATE Stock SET ";
						update += "Description='" + stock.getDescription() + "', ";
						update += "AnnualYield=" + stock.getAnnualYield() + ", ";
						update += "ClosePrice=" + stock.getClosePrice();
						update += " WHERE Ticker='" + stock.getTicker() + "';";		
						int passTest = st.executeUpdate(update);
						if (passTest==1){System.out.println("Stock " + stock.getTicker() + " updated.");}
						else{System.out.println("UPDATE SCHWAB STOCKS FAILLLL!");}
					}
					else {
						//Insert new stocks into Stock table
						String insert = "INSERT INTO Stock (Ticker, Description, AnnualYield, ClosePrice, DateAdded) VALUES (";
						insert += "'" + stock.getTicker() + "', ";
						insert += "'" + stock.getDescription() + "', ";
						insert += stock.getAnnualYield() + ", ";
						insert += stock.getClosePrice() + ", ";
						insert += "'" + convertJavaDateToSqlDate(stock.getDateAdded()).toString() + "');";
						int passTest = st.executeUpdate(insert);
						if (passTest==1){System.out.println("Stock " + stock.getTicker() + " added.");}
						else{System.out.println("INSERT SCHWAB STOCKS FAILLLL!");}
					}
					
					ResultSet schwabRes = st.executeQuery("SELECT * FROM  Schwab_Score WHERE Ticker= '" + stock.getTicker() + "';");
					if (schwabRes.next()){
						//Update stocks already in Schwab_Score table
						String update = "UPDATE Schwab_Score SET ";
						update += "SP_Rating=" + stock.getSpRating()+ ", ";
						update += "BusinessDate='" + convertJavaDateToSqlDate(stock.getBusinessDate()).toString() + "' "; 
						update += " WHERE Ticker='" + stock.getTicker() + "';";		
						int passTest = st.executeUpdate(update);
						if (passTest==1){System.out.println("Schwab score for " + stock.getTicker() + " updated.");}
						else{System.out.println("UPDATE Schwab_SCORE STOCKS FAILLLL!");}
					}
					else {
						//Insert new stocks into Schwab_Score table
						String insert = "INSERT INTO Schwab_Score (Ticker, SP_Rating, BusinessDate) VALUES (";
						insert += "'" + stock.getTicker() + "', ";
						insert += stock.getSpRating() + ", ";
						insert += "'" + convertJavaDateToSqlDate(stock.getBusinessDate()).toString() + "');";
						int passTest = st.executeUpdate(insert);
						if (passTest==1){System.out.println("Schwab score for " + stock.getTicker() + " added.");}
						else{System.out.println("INSERT SCHWAB_SCORE STOCKS FAILLLL!");}
					}
		    	}
		    	catch (Exception e){
		    		e.printStackTrace();
		    		conn.close();
		    		workbook.close();
		    	}
		    	
		    	
		    	
		    }
		    conn.close();
		    workbook.close();
			
	}
	
	public static void barronsPopulate(File fille)			//Populates database with Barrons ratings of existing stocks
			throws BiffException, IOException, WriteException, IllegalAccessException, InstantiationException, ClassNotFoundException, SQLException{		
			
			Workbook workbook = Workbook.getWorkbook(fille);			//Barrons Stocks Excel file
		    Sheet sheet = workbook.getSheet("Sheet1");
		    Class.forName(driver).newInstance();
		    Connection conn = DriverManager.getConnection(url+dbName, userName, "");		//Database connection
		    for (int i = 1; i < sheet.getRows(); i++){
		    	try{
		    		System.out.println(sheet.getCell(0, i).getContents());
			    	//A stock object is created from reading Barron's Excel sheet's data
			    	Stock stock = new Stock(sheet.getCell(0, i).getContents());
			    	try{
			    		stock.setGrade(Double.parseDouble(sheet.getCell(1, i).getContents()));
			    	}
			    	catch (NumberFormatException e){
			    		e.printStackTrace();
			    	}
			    	//DateAdded generates automatically
			    	stock.setBusinessDate(lastBusinessDate);
			    	//System.out.println(stock.toString());
			    	
			    	Statement st = conn.createStatement(); 
					ResultSet barronsRes = st.executeQuery("SELECT * FROM  Barrons_Score WHERE Ticker= '" + stock.getTicker() + "';");
					if (barronsRes.next()){
						//Update stocks already in Barrons_Score table
						System.out.println("Barrons update should be happening");
						String update = "UPDATE Barrons_Score SET ";
						update += "Grade=" + stock.getGrade() + ", ";
						update += "BusinessDate='" + convertJavaDateToSqlDate(stock.getBusinessDate()).toString() + "' "; 
						update += " WHERE Ticker='" + stock.getTicker() + "';";		
						int passTest = st.executeUpdate(update);
						if (passTest==1){System.out.println("Barrons score for " + stock.getTicker() + " updated.");}
						else{System.out.println("UPDATE Barrons_SCORE STOCKS FAILLLL!");}
					}
					else {
						//Insert new stocks into Barrons_Score table
						System.out.println("Barrons insert should be happening here");
						String insert = "INSERT INTO Barrons_Score (Ticker, Grade, BusinessDate) VALUES (";
						insert += "'" + stock.getTicker() + "', ";
						insert += stock.getGrade() + ", ";
						insert += "'" + convertJavaDateToSqlDate(stock.getBusinessDate()).toString() + "');";
						try{st.executeUpdate(insert);}
						catch (Exception e){
							//e.printStackTrace();
							System.out.println("No data for " + stock.getTicker());
						}
						//int passTest = st.executeUpdate(insert);
						//if (passTest==1){System.out.println("Barrons score for " + stock.getTicker() + " added.");}
						//else{System.out.println("INSERT BARRONS_SCORE STOCKS FAILLLL!");}
					}
		    	}
		    	catch (Exception e){
		    		e.printStackTrace();
		    		conn.close();
		    		workbook.close();
		    	}
		    	
		    	
		    	
		    }
		    conn.close();
		    workbook.close();
			
	}
	
	public static void generateResults(File fille)		//Method creates an Excel file with results of criteria filter
		throws BiffException, IOException, WriteException, IllegalAccessException, InstantiationException, ClassNotFoundException, SQLException{
		
		int outerResultCount = 0;
		int innerResultCount = 0;
		
		//Create Excel workbook and sheets
		WritableWorkbook workbook;
		workbook = Workbook.createWorkbook(fille);
		WritableSheet outerSheet = workbook.createSheet("OuterMergedStocks", 0);
		WritableSheet innerSheet = workbook.createSheet("InnerMergedStocks", 1);
		
		// Create cell formats
		WritableFont arial14font = new WritableFont(WritableFont.ARIAL, 13); 
		WritableCellFormat arial14format = new WritableCellFormat (arial14font);
		
		WritableFont boldFont = new WritableFont(WritableFont.ARIAL, 10);
	    boldFont.setBoldStyle(WritableFont.BOLD);
	    WritableCellFormat arial10bold = new WritableCellFormat(boldFont);
	    WritableCellFormat arial10boldWrapped = new WritableCellFormat(boldFont);
	    arial10boldWrapped.setWrap(true);
		
	    
	    //Add titles and other labels to worksheets
		Label titleLabel = new Label(1, 0, "Stocks to Buy (3% + Dividend Yield)", arial14format);
		outerSheet.addCell(titleLabel);
		
		Label titleLabel2 = new Label(1, 0, "Stocks to Buy (3% + Dividend Yield)", arial14format);
		innerSheet.addCell(titleLabel2);
		
		Label busDateLabel = new Label(0, 1, "Business Date: ");
		Label busDateLabel2 = new Label(0, 1, "Business Date: ");
		Date businessDate = lastBusinessDate;
		Label busDate = new Label(1, 1, businessDate.toString());		
		Label busDate2 = new Label(1, 1, businessDate.toString());
		outerSheet.addCell(busDateLabel);
		outerSheet.addCell(busDate);
		innerSheet.addCell(busDateLabel2);
		innerSheet.addCell(busDate2);
		
		Label printDateLabel = new Label(0, 2, "Print Date: ");
		Label printDateLabel2 = new Label(0, 2, "Print Date: ");
		Date printDate = new Date();
		Label printDateStr = new Label(1, 2, printDate.toString());
		Label printDateStr2 = new Label(1, 2, printDate.toString());
		outerSheet.addCell(printDateLabel);
		outerSheet.addCell(printDateStr);
		innerSheet.addCell(printDateLabel2);
		innerSheet.addCell(printDateStr2);

		
		//Add column labels to worksheets
		Label tickerLabel = new Label(0, 3, "Stock\nTicker", arial10boldWrapped);	
		Label tickerLabel2 = new Label(0, 3, "Stock\nTicker", arial10boldWrapped);	
		outerSheet.addCell(tickerLabel);
		innerSheet.addCell(tickerLabel2);
		
		Label descriptionLabel = new Label(1, 3, "Description", arial10bold);	
		Label descriptionLabel2 = new Label(1, 3, "Description", arial10bold);	 
		outerSheet.addCell(descriptionLabel);
		innerSheet.addCell(descriptionLabel2);
		
		Label yieldLabel = new Label(2, 3, "Annual\nYield", arial10boldWrapped);	 
		Label yieldLabel2 = new Label(2, 3, "Annual\nYield", arial10boldWrapped);	
		outerSheet.addCell(yieldLabel);
		innerSheet.addCell(yieldLabel2);
		
		Label priceLabel = new Label(3, 3, "Close\nPrice", arial10boldWrapped);	 
		Label priceLabel2 = new Label(3, 3, "Close\nPrice", arial10boldWrapped);	 
		outerSheet.addCell(priceLabel);
		innerSheet.addCell(priceLabel2);
		
		Label fairvalLabel = new Label(4, 3, "Fair\nValue", arial10boldWrapped);	 
		Label fairvalLabel2 = new Label(4, 3, "Fair\nValue", arial10boldWrapped);	 
		outerSheet.addCell(fairvalLabel);
		innerSheet.addCell(fairvalLabel2);
		
		Label fairvalMinusCloseLabel = new Label(5, 3, "Fair Val-\nClose", arial10boldWrapped);	 
		Label fairvalMinusCloseLabel2 = new Label(5, 3, "Fair Val-\nClose", arial10boldWrapped);	 
		outerSheet.addCell(fairvalMinusCloseLabel);
		innerSheet.addCell(fairvalMinusCloseLabel2);
		
		Label closeOverFairvalLabel = new Label(6, 3, "Close/\nFair Val", arial10boldWrapped);	 
		Label closeOverFairvalLabel2 = new Label(6, 3, "Close/\nFair Val", arial10boldWrapped);	 
		outerSheet.addCell(closeOverFairvalLabel);
		innerSheet.addCell(closeOverFairvalLabel2);
		
		Label consbuyLabel = new Label(7, 3, "Cons\nBuy Price", arial10boldWrapped);	 
		Label consbuyLabel2 = new Label(7, 3, "Cons\nBuy Price", arial10boldWrapped);
		outerSheet.addCell(consbuyLabel);
		innerSheet.addCell(consbuyLabel2);
		
		Label closeMinusConsbuyLabel = new Label(8, 3, "Close-\nCons Buy", arial10boldWrapped);	 
		Label closeMinusConsbuyLabel2 = new Label(8, 3, "Close-\nCons Buy", arial10boldWrapped);
		outerSheet.addCell(closeMinusConsbuyLabel);
		innerSheet.addCell(closeMinusConsbuyLabel2);
		
		Label sectorLabel = new Label(9, 3, "Sector", arial10bold);	
		Label sectorLabel2 = new Label(9, 3, "Sector", arial10bold);	
		outerSheet.addCell(sectorLabel);
		innerSheet.addCell(sectorLabel2);
		
		//CellView cv = innerSheet.getColumnView(4);
        //cv.setSize(10 * 256 + 100); /* Every character is 256 units wide, so scale it. */
        //innerSheet.setColumnView(4, cv);
		
		Label industryLabel = new Label(10, 3, "Industry", arial10bold);	
		Label industryLabel2 = new Label(10, 3, "Industry", arial10bold);	 
		outerSheet.addCell(industryLabel);
		innerSheet.addCell(industryLabel2);
		
		Label marketCap = new Label(11, 3, "Market Cap ($mil)", arial10bold);
		Label marketCap2 = new Label(11, 3, "Market Cap ($mil)", arial10bold);
		outerSheet.addCell(marketCap);
		innerSheet.addCell(marketCap2);
		
		Label mstarLabel = new Label(12, 3, "Mstar\nRating", arial10boldWrapped);	
		Label mstarLabel2 = new Label(12, 3, "Mstar\nRating", arial10boldWrapped);
		outerSheet.addCell(mstarLabel);
		innerSheet.addCell(mstarLabel2);
		
		Label schwabLabel = new Label(13, 3, "Schwab S&P\nRating", arial10boldWrapped);	 
		Label schwabLabel2 = new Label(13, 3, "Schwab S&P\nRating", arial10boldWrapped);	
		outerSheet.addCell(schwabLabel);
		innerSheet.addCell(schwabLabel2);
		
		Label gradeLabel = new Label(14, 3, "Barrons\nGrade", arial10boldWrapped);	 
		Label gradeLabel2 = new Label(14, 3, "Barrons\nGrade", arial10boldWrapped);
		outerSheet.addCell(gradeLabel);
		innerSheet.addCell(gradeLabel2);
		
		
		sheetAutoFitColumns(outerSheet);			//May need to change this method so columns aren't as wide as headers
		sheetAutoFitColumns(innerSheet);
		
		//Cell sizing
		
		CellView tickerCV = innerSheet.getColumnView(0);
        tickerCV.setSize(7 * 256 + 100); 
        innerSheet.setColumnView(0, tickerCV);
        
        CellView descriptionCV = innerSheet.getColumnView(1);
        descriptionCV.setSize(27 * 256 + 100); 
        innerSheet.setColumnView(1, descriptionCV);
        
        CellView yieldCV = innerSheet.getColumnView(2);
        yieldCV.setSize(8 * 256 + 100); 
        innerSheet.setColumnView(2, yieldCV);
        
        CellView priceCV = innerSheet.getColumnView(3);
        priceCV.setSize(8 * 256 + 100); 
        innerSheet.setColumnView(3, priceCV);
        
        CellView fairvalCV = innerSheet.getColumnView(4);
        fairvalCV.setSize(8 * 256 + 100); 
        innerSheet.setColumnView(4, fairvalCV);
        
        CellView fairvalMinusCloseCV = innerSheet.getColumnView(5);
        fairvalMinusCloseCV.setSize(8 * 256 + 100); 
        innerSheet.setColumnView(5, fairvalMinusCloseCV);
        
        CellView closeOverFairvalCV = innerSheet.getColumnView(6);
        closeOverFairvalCV.setSize(8 * 256 + 100); 
        innerSheet.setColumnView(6, closeOverFairvalCV);
		
        CellView consbuyCV = innerSheet.getColumnView(7);
        consbuyCV.setSize(8 * 256 + 100); 
        innerSheet.setColumnView(7, consbuyCV);
        
        CellView closeMinusConsbuyCV = innerSheet.getColumnView(8);
        closeMinusConsbuyCV.setSize(8 * 256 + 100); 
        innerSheet.setColumnView(8, closeMinusConsbuyCV);
		
		CellView sectorCV = innerSheet.getColumnView(9);
        sectorCV.setSize(21 * 256 + 100); 
        innerSheet.setColumnView(9, sectorCV);
        
        CellView industryCV = innerSheet.getColumnView(10);
        industryCV.setSize(33 * 256 + 100); 
        innerSheet.setColumnView(10, industryCV);
        
        //MarketCap 11
        
        CellView mstarCV = innerSheet.getColumnView(12);
        mstarCV.setSize(8 * 256 + 100); 
        innerSheet.setColumnView(12, mstarCV);
        
        CellView spCV = innerSheet.getColumnView(13);
        spCV.setSize(9 * 256 + 100); 
        innerSheet.setColumnView(13, spCV);
        
        CellView barronsCV = innerSheet.getColumnView(14);
        barronsCV.setSize(8 * 256 + 100); 
        innerSheet.setColumnView(14, barronsCV);
		
	    Class.forName(driver).newInstance();
	    Connection conn = DriverManager.getConnection(url+dbName, userName, "");		//Database connection
	    
	    //Run query to obtain stocks where dividend yield is >= 3.00% and pass either condition:
	    //1) Have a Morningstar rating of 4 or 5 stars and have a Fair Value Estimate >= .95*Consider Buying Price
	    //2) Have a Schwab S&P Rating of 4 or 5 stars
	    Statement st = conn.createStatement();
	    String outerMergeSelect = "SELECT S.Ticker, S.Description, S.AnnualYield, S.ClosePrice, ";
	    outerMergeSelect += "M.StarRating, M.FairValue, M.ConsiderBuyPrice, ";
	    outerMergeSelect += "Sc.SP_Rating ";
	    String outerMergeFrom = "FROM Stock S LEFT JOIN Morningstar_Score M ON S.Ticker = M.Ticker ";
	    outerMergeFrom += "LEFT JOIN Schwab_Score Sc ON S.Ticker = Sc.Ticker ";
	    String outerMergeWhere = "WHERE S.AnnualYield >= 3.00 AND ";
	    outerMergeWhere += "((M.StarRating >= 4 AND M.FairValue >= M.ConsiderBuyPrice*.95) OR Sc.SP_Rating >= 4);";
		ResultSet res = st.executeQuery(outerMergeSelect + outerMergeFrom + outerMergeWhere);
		
		//Add results to Outer Merge Sheet
		while (res.next()){
			Label tempTicker = new Label(0, outerResultCount+4, res.getString(1));
			outerSheet.addCell(tempTicker);
			
			Label tempDescription = new Label(1, outerResultCount+4, res.getString(2));
			outerSheet.addCell(tempDescription);
			
			Number tempYield = new Number(2, outerResultCount+4, res.getDouble(3)/100.00);
			tempYield.setCellFormat(new WritableCellFormat(NumberFormats.PERCENT_FLOAT));
			outerSheet.addCell(tempYield);
			
			Number tempPrice = new Number(3, outerResultCount+4, res.getDouble(4));
			tempPrice.setCellFormat(new WritableCellFormat(NumberFormats.ACCOUNTING_FLOAT));
			outerSheet.addCell(tempPrice);
			
			
			if (res.getInt(5) > 0){
				Number tempMSRating = new Number(4, outerResultCount+4, res.getInt(5));
				outerSheet.addCell(tempMSRating);
				
				Number tempFairVal = new Number(5, outerResultCount+4, res.getDouble(6));
				tempFairVal.setCellFormat(new WritableCellFormat(NumberFormats.ACCOUNTING_FLOAT));
				outerSheet.addCell(tempFairVal);
				
				Number tempConBuy = new Number(6, outerResultCount+4, res.getDouble(7));
				tempConBuy.setCellFormat(new WritableCellFormat(NumberFormats.ACCOUNTING_FLOAT));
				outerSheet.addCell(tempConBuy);
			}
			
			if (res.getInt(8) > 0){
				Number tempSPRating = new Number(7, outerResultCount+4, res.getInt(8));
				outerSheet.addCell(tempSPRating);
			}
			
			System.out.println(res.getString(1));
			outerResultCount++;
		}
		
		//Run query to obtain stocks where dividend yield is >= 3.00% and pass both conditions:
	    //1) Have a Morningstar rating of 4 or 5 stars and have a Fair Value Estimate >= .95*Consider Buying Price
	    //2) Have a Schwab S&P Rating of 4 or 5 stars
	    Statement st2 = conn.createStatement();
	    String innerMergeSelect = "SELECT S.Ticker, S.Description, S.AnnualYield, S.ClosePrice, ";
	    innerMergeSelect += "S.Sector, S.Industry, S.MarketCap, ";
	    innerMergeSelect += "M.StarRating, M.FairValue, M.ConsiderBuyPrice, ";
	    innerMergeSelect += "Sc.SP_Rating, ";
	    innerMergeSelect += "B.Grade ";
	    String innerMergeFrom = "FROM Stock S LEFT JOIN Morningstar_Score M ON S.Ticker = M.Ticker ";
	    innerMergeFrom += "LEFT JOIN Schwab_Score Sc ON S.Ticker = Sc.Ticker ";
	    innerMergeFrom += "LEFT JOIN Barrons_Score B ON S.Ticker = B.Ticker ";
	    String innerMergeWhere = "WHERE S.AnnualYield >= 3.00 AND ";
	    innerMergeWhere += "((M.StarRating >= 4 AND M.FairValue >= M.ConsiderBuyPrice*.95) AND Sc.SP_Rating >= 4) ";
	    innerMergeWhere += "AND M.BusinessDate='" + convertJavaDateToSqlDate(lastBusinessDate).toString() + "' ";
	    innerMergeWhere += "AND Sc.BusinessDate='" + convertJavaDateToSqlDate(lastBusinessDate).toString() + "';";
		ResultSet res2 = st2.executeQuery(innerMergeSelect + innerMergeFrom + innerMergeWhere);
		
		//Add results to Inner Merge Sheet
		while (res2.next()){
			
			Label tempTicker = new Label(0, innerResultCount+4, res2.getString(1));
			innerSheet.addCell(tempTicker);
			
			Label tempDescription = new Label(1, innerResultCount+4, res2.getString(2));
			innerSheet.addCell(tempDescription);
			
			Number tempYield = new Number(2, innerResultCount+4, res2.getDouble(3)/100.00);
			tempYield.setCellFormat(new WritableCellFormat(NumberFormats.PERCENT_FLOAT));
			innerSheet.addCell(tempYield);
			
			Number tempPrice = new Number(3, innerResultCount+4, res2.getDouble(4));
			tempPrice.setCellFormat(new WritableCellFormat(NumberFormats.ACCOUNTING_FLOAT));
			innerSheet.addCell(tempPrice);
			
			Number tempFairVal = new Number(4, innerResultCount+4, res2.getDouble(9));
			tempFairVal.setCellFormat(new WritableCellFormat(NumberFormats.ACCOUNTING_FLOAT));
			innerSheet.addCell(tempFairVal);
			
			Number tempFairValMinusPrice = new Number(5, innerResultCount+4, (res2.getDouble(9) - res2.getDouble(4)));
			tempFairValMinusPrice.setCellFormat(new WritableCellFormat(NumberFormats.ACCOUNTING_FLOAT));
			innerSheet.addCell(tempFairValMinusPrice);
			
			Number tempPriceOverFairVal = new Number(6, innerResultCount+4, (res2.getDouble(4)/res2.getDouble(9)));
			tempPriceOverFairVal.setCellFormat(new WritableCellFormat(NumberFormats.PERCENT_FLOAT));
			innerSheet.addCell(tempPriceOverFairVal);
			
			Number tempConBuy = new Number(7, innerResultCount+4, res2.getDouble(10));
			tempConBuy.setCellFormat(new WritableCellFormat(NumberFormats.ACCOUNTING_FLOAT));
			innerSheet.addCell(tempConBuy);
			
			Number tempPriceMinusConBuy = new Number(8, innerResultCount+4, (res2.getDouble(4)-res2.getDouble(10)));
			tempPriceMinusConBuy.setCellFormat(new WritableCellFormat(NumberFormats.ACCOUNTING_FLOAT));
			innerSheet.addCell(tempPriceMinusConBuy);
			
			Label tempSector = new Label(9, innerResultCount+4, res2.getString(5));
			innerSheet.addCell(tempSector);
			
			Label tempIndustry = new Label(10, innerResultCount+4, res2.getString(6));
			innerSheet.addCell(tempIndustry);
			
			Number tempCap = new Number(11, innerResultCount+4, res2.getDouble(7));
			tempCap.setCellFormat(new WritableCellFormat(NumberFormats.ACCOUNTING_FLOAT));
			innerSheet.addCell(tempCap);
			
			Number tempMSRating = new Number(12, innerResultCount+4, res2.getInt(8));
			innerSheet.addCell(tempMSRating);
			
			Number tempSPRating = new Number(13, innerResultCount+4, res2.getInt(11));
			innerSheet.addCell(tempSPRating);
			
			if (res2.getDouble(12)!=0){
				Number tempGrade = new Number(14, innerResultCount+4, res2.getDouble(12));
				innerSheet.addCell(tempGrade);
			}
			else {
				Blank tempGrade = new Blank(14, innerResultCount+4);
				innerSheet.addCell(tempGrade);
			}
			
			System.out.println(res2.getString(1));
			innerResultCount++;
		}
		

		conn.close();
		workbook.moveSheet(1, 0);
		workbook.write();
		workbook.close();
	}
	
	private static void sheetAutoFitColumns(WritableSheet sheet) {
	    for (int i = 0; i < sheet.getColumns(); i++) {
	        Cell[] cells = sheet.getColumn(i);
	        int longestStrLen = -1;

	        if (cells.length == 0)
	            continue;

	        /* Find the widest cell in the column. */
	        for (int j = 0; j < cells.length; j++) {
	            if ( cells[j].getContents().length() > longestStrLen ) {
	                String str = cells[j].getContents();
	                if (str == null || str.isEmpty())
	                    continue;
	                longestStrLen = str.trim().length();
	            }
	        }

	        /* If not found, skip the column. */
	        if (longestStrLen == -1) 
	            continue;

	        /* If wider than the max width, crop width */
	        if (longestStrLen > 255)
	            longestStrLen = 255;

	        CellView cv = sheet.getColumnView(i);
	        cv.setSize(longestStrLen * 256 + 100); /* Every character is 256 units wide, so scale it. */
	        sheet.setColumnView(i, cv);
	    }
	}
	
	public static void addRowView(WritableSheet sheet, int row, Colour colour)
		throws WriteException {
	    
	    WritableCellFormat cellFormat = new WritableCellFormat();
	    cellFormat.setBackground(colour);
	    //cellFormat.setBorder(Border.ALL, BorderLineStyle.THIN);
	    
	    CellView cellView = new CellView();
	    cellView.setFormat(cellFormat);
	    sheet.setRowView(row, cellView);
	}

	public static double round(double value, int places) {
	    if (places < 0) throw new IllegalArgumentException();

	    java.math.BigDecimal bd = new java.math.BigDecimal(value);
	    bd = bd.setScale(places, java.math.RoundingMode.HALF_UP);
	    return bd.doubleValue();
	}
	
	public static java.sql.Date convertJavaDateToSqlDate(java.util.Date date) {
	    return new java.sql.Date(date.getTime());
	}
	
	public static java.util.Date convertSqlDateToJavaDate(java.sql.Date date) {
	    return new java.util.Date(date.getTime());
	}
	
}
