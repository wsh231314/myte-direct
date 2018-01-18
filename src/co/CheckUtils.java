package co;

import java.sql.Date;
import java.text.SimpleDateFormat;

public class CheckUtils {
	
	/**
	 * check the value is empty or not
	 * 
	 * @param strValue
	 * @return
	 */
	public static boolean isEmptyString (String strValue) {
		boolean result = false;
		
		if (strValue == null || strValue.trim().equals("")) {
			result = true;
		}
		
		return result;
	}
	
	/**
	 * check the value is number or not
	 * 
	 * @param strValue
	 * @return
	 */
	public static boolean isNumber (String strValue) {
		boolean result = false;
		
		try {
			Double.parseDouble(strValue);
			result = true;
		} catch (Exception e) {
			result = false;
		}
		
		return result;
	}
	
	/**
	 * check the value is date or not
	 * 
	 * @param strValue
	 * @return
	 */
	public static boolean isDateYYYYMMDD (String strValue) {
		boolean result = false;
		
		try {
			SimpleDateFormat sdf = new SimpleDateFormat("yyyy/MM/dd");
			sdf.parse(strValue);
			result = true;
		} catch (Exception e) {
			result = false;
		}
		
		return result;
	}

}
