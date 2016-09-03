import java.io.File;
import java.io.IOException;
import java.util.Date;

//Stock class includes all website ratings, unlike database where ratings
//are in separate tables. May change structure

public class Stock {

		private String ticker;
		private String description;
		private double annualYield;
		private double closePrice;
		private String sector;
		private String industry;
		private double marketCap;
		private Date dateAdded;		//May change type later, currently refers to SQL date
		private double grade;
		private double sentiment;
		private Date businessDate;
		private short starRating;
		private double fairValue;
		private double considerBuyingPrice;
		private short spRating;
		
		public Stock(String tick){
			ticker = tick;
		}
		
		public Stock(String tick, String descr){
			setTicker(tick);
			setDescription(descr);
			dateAdded = new Date();
		}
		
		public Stock(String tick, String descr, double yield, double price){
			setTicker(tick);
			setDescription(descr);
			annualYield = yield;
			closePrice = price;
			dateAdded = new Date();
		}

		public String getTicker() {
			return ticker;
		}

		public void setTicker(String ticker) {
			this.ticker = ticker;
		}

		public String getDescription() {
			return description;
		}

		public void setDescription(String description) {
			this.description = description;
		}
		
		public double getAnnualYield() {
			return annualYield;
		}

		public void setAnnualYield(double annualYield) {
			this.annualYield = annualYield;
		}

		public double getClosePrice() {
			return closePrice;
		}

		public void setClosePrice(double closePrice) {
			this.closePrice = closePrice;
		}
		
		public String getSector() {
			return sector;
		}

		public void setSector(String sector) {
			this.sector = sector;
		}
		
		public String getIndustry() {
			return industry;
		}

		public void setIndustry(String industry) {
			this.industry = industry;
		}

		public double getMarketCap() {
			return marketCap;
		}

		public void setMarketCap(double marketCap) {
			this.marketCap = marketCap;
		}

		public Date getDateAdded() {
			return dateAdded;
		}

		public void setDateAdded(Date dateAdded) {
			this.dateAdded = dateAdded;
		}

		public double getGrade() {
			return grade;
		}

		public void setGrade(double grade) {
			this.grade = grade;
		}

		public double getSentiment() {
			return sentiment;
		}

		public void setSentiment(double sentiment) {
			this.sentiment = sentiment;
		}

		public Date getBusinessDate() {
			return businessDate;
		}

		public void setBusinessDate(Date businessDate) {
			this.businessDate = businessDate;
		}

		public short getStarRating() {
			return starRating;
		}

		public void setStarRating(short starRating) {
			this.starRating = starRating;
		}

		public double getFairValue() {
			return fairValue;
		}

		public void setFairValue(double fairValue) {
			this.fairValue = fairValue;
		}

		public double getConsiderBuyingPrice() {
			return considerBuyingPrice;
		}

		public void setConsiderBuyingPrice(double considerBuyingPrice) {
			this.considerBuyingPrice = considerBuyingPrice;
		}

		public short getSpRating() {
			return spRating;
		}

		public void setSpRating(short spRating) {
			this.spRating = spRating;
		}
		
		public String toString(){
			
			String result = "";
			result = ticker + "\t" + description + "\t" + annualYield  + "% \t $" + closePrice + "\t";
			result += dateAdded.toString() + "\n";
			result += sector + "\t "+ industry + "\t" + marketCap + "\n";
			result += "Morningstar Data: \t" + starRating + "\t" + fairValue + "\t" + considerBuyingPrice + "\n";
			result += "Schwab Data: \t" + spRating + "\n" ;
			result += "Business Date: \t" + businessDate.toString();
			
			return result;
		}
}
