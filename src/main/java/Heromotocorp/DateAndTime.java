package Heromotocorp;


import org.joda.time.DateTime;
public class DateAndTime {

	static DateTime now = new DateTime();
	static String month= now.monthOfYear().getAsText();
	static int day = now.getDayOfMonth();
	static int min=now.getHourOfDay();
	static int second = now.getMinuteOfHour();
	
	
	public static String  customized_time(){
		String time_curr =month+"-"+day+" :"+min+":"+second;
		return time_curr;
	}
	public static String  customized_Date(){
		String time_date =month+"-"+day+" "+min;
		return time_date;
	}

}
