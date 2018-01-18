package co;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

import javax.swing.JFileChooser;
import javax.swing.JOptionPane;

import org.apache.log4j.Logger;
import org.apache.log4j.PropertyConfigurator;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.NoSuchElementException;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.openqa.selenium.support.ui.Select;

public class MyteDirect {
	
	static Logger logger = null;
	
	public static void main(String[] args) {
		
		PropertyConfigurator.configure(MyteDirect.class.getClassLoader().getResourceAsStream("log4j.properties"));
		logger = Logger.getLogger(MyteDirect.class);
		
		System.setProperty("webdriver.ie.driver", "./driver/IEDriverServer.exe");
		InternetExplorerDriver driver = new InternetExplorerDriver();

		try {
			// 
			driver.get("https://myte.accenture.com");

			By timeId = By.id("ctl00_ctl00_MainContentPlaceHolder_Time");

			boolean blnSearch = true;
			int iWaitCount = 50;

			// wait for myte to open, until 50 seconds past.
			while (blnSearch && iWaitCount > 0) {
				try {
					driver.findElement(timeId);
					blnSearch = false;
				} catch (Exception e) {
					iWaitCount = iWaitCount - 1;
					try {
						Thread.sleep(1000);
					} catch (InterruptedException e1) {
						logger.error("load myte has error!", e1);
					}
				}
			}

			// click the expense button
			WebElement expenseElement = driver.findElementByXPath("//a[contains(@href, 'ExpenseSheetPage.aspx')]");
			expenseElement.click();


			// wait for expense input page to open, until 20 seconds past.
			By expenseAdd = By.id("ctl00_ctl00_MainContentPlaceHolder_ContentPlaceHolder_TimeReport_AddExpenseList_new_toggleDiv");
			blnSearch = true;
			iWaitCount = 50;

			WebElement expenseAddelement = null;
			while (blnSearch && iWaitCount > 0) {
				try {
					expenseAddelement = driver.findElement(expenseAdd);
					blnSearch = false;
				} catch (Exception e) {
					iWaitCount = iWaitCount - 1;
					try {
						Thread.sleep(1000);
					} catch (InterruptedException e1) {
						logger.error("load expense page has error!", e1);
					}
				}
			}

			// if the expense has been locked, show the message and exit
			if (expenseAddelement != null) {
				boolean disabled = expenseAddelement.getAttribute("class").equals("disabled");
				if (disabled) {
					JOptionPane.showMessageDialog(null, "Timesheet has been locked!");
					logger.info("Timesheet has been locked!");
					System.exit(0);
				}
			}

			// show the file select pane
			File fileExpenseFile = showAndRemindeInput();

			if (fileExpenseFile != null) {
				String strMessage = checkSelectedFile(fileExpenseFile);

				if (!strMessage.equals("")) {
					JOptionPane.showMessageDialog(null, strMessage);

					fileExpenseFile = directInputOrExit() ;
				}
			} else {
				fileExpenseFile = directInputOrExit() ;
			}

			// Taxi
			XSSFWorkbook book = new XSSFWorkbook(new FileInputStream(fileExpenseFile));
			XSSFSheet taxiSheet = book.getSheetAt(0);

			boolean blnRead = true;
			boolean rowResult = false;
			int iRow = 2;
			while (blnRead) {
				
				if (iRow != 2) {
					XSSFRow preRow = taxiSheet.getRow(iRow - 1);
					
					// result check
					By popupId = By.id("ctl00_ctl00_MainContentPlaceHolder_ContentPlaceHolder_TimeReport_pnl_ExpenseDetails");
					boolean blnResult = false;
					WebElement popElement = null; 
					try {
						popElement = driver.findElement(popupId);
					} catch (NoSuchElementException e) {
						blnResult = true;
					}
					String strResult = "";
					if ((blnResult || !popElement.isDisplayed()) && rowResult) {
						strResult = "Input successed!";
					} else {
						strResult = "Input failed! check the data please!";
						
						if (!blnResult) {
							By cancelId = By.id("ctl00_ctl00_MainContentPlaceHolder_ContentPlaceHolder_TimeReport_btn_DetailCancel_bottom");
							WebElement cancelElement = driver.findElement(cancelId);
							cancelElement.click();
							
							// wait 1 second
							try {
								Thread.sleep(3000);
							} catch (InterruptedException e) {
								e.printStackTrace();
							}
						}
						
					}
					
					setValueToCell(preRow, 8, strResult);
				}
				
				rowResult = false;

				XSSFRow row = taxiSheet.getRow(iRow);
				//No
				String strNo = getStringFromCell(row.getCell(1));
				//WBS
				String WBS = getStringFromCell(row.getCell(2));
				//Amount
				String strAmount = getStringFromCell(row.getCell(3));
				//On
				String strOn = getStringFromCell(row.getCell(4));
				//From
				String strFrom = getStringFromCell(row.getCell(5));
				//To
				String strTo = getStringFromCell(row.getCell(6));
				//Reason
				String strReason = getStringFromCell(row.getCell(7));
				
				iRow++;

				// Check
				if (CheckUtils.isEmptyString(strNo)) {
					break;
				}

				// number check
				if (!CheckUtils.isNumber(strAmount)) {
					continue;
				}

				// date check
				if (!CheckUtils.isDateYYYYMMDD(strOn)) {
					continue;
				}

				// open the taxi expense page

				By expenseDiv = By.id("ctl00_ctl00_MainContentPlaceHolder_ContentPlaceHolder_TimeReport_AddExpenseList_new_toggleDiv");

				driver.findElement(expenseDiv).click();

				// loop the data and input myte
				By expenseSelect = By.xpath("//div[@class='selectToList listExpenses']/ul[@class='options']/li[contains(@rel, 'EX01')]");
				driver.findElement(expenseSelect).click();

				// input the taxi and save
				// WBS
				By wbsId = By.id("ctl00_ctl00_MainContentPlaceHolder_ContentPlaceHolder_TimeReport_expense_DetailsControl_expenseProjectDropDown");
				Select wbsSelect = new Select(driver.findElement(wbsId));
				if (CheckUtils.isEmptyString(WBS)) {
					wbsSelect.selectByIndex(1);
					logger.info("No WBS code, choose the first option!");
				} else {
					try {
						wbsSelect.selectByValue(WBS);
					} catch (Exception e) {
						wbsSelect.selectByIndex(1);
						logger.info("No WBS code match, choose the first option!");
					}
				}

				// strAmount
				By amountId = By.id("ctl00_ctl00_MainContentPlaceHolder_ContentPlaceHolder_TimeReport_expense_DetailsControl_8467");
				WebElement amountElement = driver.findElement(amountId);
				amountElement.clear();
				amountElement.sendKeys(strAmount);

				// On
				By onId = By.id("ctl00_ctl00_MainContentPlaceHolder_ContentPlaceHolder_TimeReport_expense_DetailsControl_expenseDate");
				WebElement onElement = driver.findElement(onId);
				onElement.clear();
				onElement.sendKeys(strOn);

				// From
				By fromId = By.id("ctl00_ctl00_MainContentPlaceHolder_ContentPlaceHolder_TimeReport_expense_DetailsControl_8498");
				try {
					WebElement fromElement = driver.findElement(fromId);
					fromElement.clear();
					fromElement.sendKeys(strFrom);
				} catch (Exception e) {
					logger.info("Do not have the [From] input field, ignore it!");
				}

				// To
				By toId = By.id("ctl00_ctl00_MainContentPlaceHolder_ContentPlaceHolder_TimeReport_expense_DetailsControl_8499");
				try {
					WebElement fromElement = driver.findElement(toId);
					fromElement.clear();
					fromElement.sendKeys(strTo);
				} catch (Exception e) {
					logger.info("Do not have the [To] input field, ignore it!");
				}

				// Reason
				By reasonId = By.id("ctl00_ctl00_MainContentPlaceHolder_ContentPlaceHolder_TimeReport_expense_DetailsControl_8471");
				Select reasonSelect = new Select(driver.findElement(reasonId));
				try {
					reasonSelect.selectByValue(strReason);
				} catch (Exception e) {
					reasonSelect.selectByIndex(1);
					logger.info("No Reason code match, choose the first option!");
				}

				// ok button
				By okButtonId = By.id("ctl00_ctl00_MainContentPlaceHolder_ContentPlaceHolder_TimeReport_btn_DetailOK_bottom");
				driver.findElement(okButtonId).click();

				// wait 1 second
				try {
					Thread.sleep(3000);
				} catch (InterruptedException e) {
					e.printStackTrace();
				}
				
				rowResult = true;
			}
			
			// meals
			XSSFSheet mealsSheet = book.getSheetAt(1);

			blnRead = true;
			rowResult = false;
			iRow = 2;
			while (blnRead) {
				
				if (iRow != 2) {
					XSSFRow preRow = mealsSheet.getRow(iRow - 1);
					
					// result check
					By popupId = By.id("ctl00_ctl00_MainContentPlaceHolder_ContentPlaceHolder_TimeReport_pnl_ExpenseDetails");
					boolean blnResult = false;
					WebElement popElement = null; 
					try {
						popElement = driver.findElement(popupId);
					} catch (NoSuchElementException e) {
						blnResult = true;
					}
					String strResult = "";
					if ((blnResult || !popElement.isDisplayed()) && rowResult) {
						strResult = "Input successed!";
					} else {
						strResult = "Input failed! check the data please!";
						
						if (!blnResult) {
							By cancelId = By.id("ctl00_ctl00_MainContentPlaceHolder_ContentPlaceHolder_TimeReport_btn_DetailCancel_bottom");
							WebElement cancelElement = driver.findElement(cancelId);
							cancelElement.click();
							
							// wait 1 second
							try {
								Thread.sleep(3000);
							} catch (InterruptedException e) {
								e.printStackTrace();
							}
						}
					}
					
					setValueToCell(preRow, 9, strResult);
				}
				
				rowResult = false;
				
				XSSFRow row = mealsSheet.getRow(iRow);

				//No
				String strNo = getStringFromCell(row.getCell(1));
				//WBS
				String WBS = getStringFromCell(row.getCell(2));
				//Amount
				String strAmount = getStringFromCell(row.getCell(3));
				//On
				String strOn = getStringFromCell(row.getCell(4));
				//Reason
				String strReason = getStringFromCell(row.getCell(5));
				//Restaurant
				String strRestaurant = getStringFromCell(row.getCell(6));
				//Number of Attendees
				String strNumAttendees = getStringFromCell(row.getCell(7));
				//Internal attendees
				String strAttendees = getStringFromCell(row.getCell(8));
				
				iRow++;

				// Check
				if (CheckUtils.isEmptyString(strNo)) {
					break;
				}

				// number check
				if (!CheckUtils.isNumber(strAmount)) {
					continue;
				}

				// date check
				if (!CheckUtils.isDateYYYYMMDD(strOn)) {
					continue;
				}
				
				// number check
				if (!CheckUtils.isNumber(strNumAttendees)) {
					continue;
				}

				// open the taxi expense page

				By expenseDiv = By.id("ctl00_ctl00_MainContentPlaceHolder_ContentPlaceHolder_TimeReport_AddExpenseList_new_toggleDiv");

				driver.findElement(expenseDiv).click();

				// loop the data and input myte
				By expenseSelect = By.xpath("//div[@class='selectToList listExpenses']/ul[@class='options']/li[contains(@rel, 'EX04')]");
				driver.findElement(expenseSelect).click();

				// input the taxi and save
				// WBS
				By wbsId = By.id("ctl00_ctl00_MainContentPlaceHolder_ContentPlaceHolder_TimeReport_expense_DetailsControl_expenseProjectDropDown");
				Select wbsSelect = new Select(driver.findElement(wbsId));
				if (CheckUtils.isEmptyString(WBS)) {
					wbsSelect.selectByIndex(1);
					logger.info("No WBS code, choose the first option!");
				} else {
					try {
						wbsSelect.selectByValue(WBS);
					} catch (Exception e) {
						wbsSelect.selectByIndex(1);
						logger.info("No WBS code match, choose the first option!");
					}
				}

				// strAmount
				By amountId = By.id("ctl00_ctl00_MainContentPlaceHolder_ContentPlaceHolder_TimeReport_expense_DetailsControl_8779");
				WebElement amountElement = driver.findElement(amountId);
				amountElement.clear();
				amountElement.sendKeys(strAmount);

				// On
				By onId = By.id("ctl00_ctl00_MainContentPlaceHolder_ContentPlaceHolder_TimeReport_expense_DetailsControl_expenseDate");
				WebElement onElement = driver.findElement(onId);
				onElement.clear();
				onElement.sendKeys(strOn);

				// Reason
				By reasonId = By.id("ctl00_ctl00_MainContentPlaceHolder_ContentPlaceHolder_TimeReport_expense_DetailsControl_8783");
				Select reasonSelect = new Select(driver.findElement(reasonId));
				try {
					reasonSelect.selectByValue(strReason);
				} catch (Exception e) {
					reasonSelect.selectByIndex(1);
					logger.info("No Reason code match, choose the first option!");
				}
				
				// Restaurant
				By restaurantId = By.id("ctl00_ctl00_MainContentPlaceHolder_ContentPlaceHolder_TimeReport_expense_DetailsControl_8790");
				WebElement restaurantElement = driver.findElement(restaurantId);
				restaurantElement.clear();
				restaurantElement.sendKeys(strRestaurant);
				
				// Number of Attendees
				By numAttendeesId = By.id("ctl00_ctl00_MainContentPlaceHolder_ContentPlaceHolder_TimeReport_expense_DetailsControl_8794");
				WebElement numAttendeesElement = driver.findElement(numAttendeesId);
				numAttendeesElement.clear();
				numAttendeesElement.sendKeys(strNumAttendees);
				
				// Internal attendees
				By attendeesId = By.id("ctl00_ctl00_MainContentPlaceHolder_ContentPlaceHolder_TimeReport_expense_DetailsControl_8801");
				WebElement attendeesElement = driver.findElement(attendeesId);
				attendeesElement.clear();
				attendeesElement.sendKeys(strAttendees);
				
				// pre-approval
				By preApprovalId = By.id("ctl00_ctl00_MainContentPlaceHolder_ContentPlaceHolder_TimeReport_expense_DetailsControl_8823");
				WebElement preApprovalElement = driver.findElement(preApprovalId);
				if (!preApprovalElement.isSelected()) {
					preApprovalElement.click();
				}

				// ok button
				By okButtonId = By.id("ctl00_ctl00_MainContentPlaceHolder_ContentPlaceHolder_TimeReport_btn_DetailOK_bottom");
				driver.findElement(okButtonId).click();

				// wait 1 second
				try {
					Thread.sleep(3000);
				} catch (InterruptedException e) {
					e.printStackTrace();
				}
				rowResult = true;
			}
			
			// set output file name
			
			String strFileName = fileExpenseFile.getPath();
			strFileName = strFileName.substring(0, strFileName.lastIndexOf(".")).concat("_result.xlsx");
			File outputFile = new File(strFileName);
			
			book.write(new FileOutputStream(outputFile));

		} catch (Exception e) {
			logger.error("It has something wrong when input the expnese!", e);
		} finally {
			driver.quit();;
		}

	}
	
	/**
	 * remind to select the file
	 * 
	 * @return
	 */
	private static File directInputOrExit() {
		
		File fileExpenseFile = null;
		
		int iResult = JOptionPane.showConfirmDialog(null, "You did not select the right expense file. If you want to exit, please select yes.");
		
		if (iResult != JOptionPane.YES_OPTION) {
			fileExpenseFile = showAndRemindeInput();
			
			if (fileExpenseFile != null) {
				
				String strMessage = checkSelectedFile(fileExpenseFile);
				
				if (!strMessage.equals("")) {
					JOptionPane.showMessageDialog(null, strMessage);
					
					fileExpenseFile = directInputOrExit() ;
				} 
			}
			
		} else {
			logger.info("You did not select the right expense file. you have choose to exit!");
			System.exit(1);
		}
		
		return fileExpenseFile;
	}
	
	/**
	 * show the file select dialog and return the result
	 * 
	 * @return
	 */
	private static File showAndRemindeInput() {
		JFileChooser fileChooser = new JFileChooser();
		fileChooser.setDialogTitle("Please select the myte expense file which download from auto OCR page!");
		logger.info("Please select the myte expense file which download from auto OCR page!");
		fileChooser.showOpenDialog(null);
		File fileExpenseFile = fileChooser.getSelectedFile();
		return fileExpenseFile;
	}
	
	/**
	 * check the input file
	 * 
	 * @param inputfile
	 * @return
	 */
	private static String checkSelectedFile(File inputfile) {
		String strMessage = "";
		
		// file input 
		if (!inputfile.getName().toLowerCase().endsWith(".xlsx")) {
			strMessage = "The expense file name is not [xlsx] file!";
			return strMessage;
		}
		
		XSSFWorkbook book = null;
		try {
			book = new XSSFWorkbook(new FileInputStream(inputfile));
			int iCount = book.getNumberOfSheets();
			
			if (iCount != 2) {
				strMessage = "The selected file has wrong format!";
				return strMessage;
			}
			
			// taxi sheet
			XSSFSheet taxiSheet = book.getSheetAt(0);
			XSSFRow taxiRow = taxiSheet.getRow(1);
			
			if (!taxiRow.getCell(1).getStringCellValue().equals("No.")
					|| !taxiRow.getCell(2).getStringCellValue().equals("WBS")
					|| !taxiRow.getCell(3).getStringCellValue().equals("Amount")
					|| !taxiRow.getCell(4).getStringCellValue().equals("On")
					|| !taxiRow.getCell(5).getStringCellValue().equals("From")
					|| !taxiRow.getCell(6).getStringCellValue().equals("To")
					|| !taxiRow.getCell(7).getStringCellValue().equals("Reason")) {
				strMessage = "The selected file has wrong format int the taxi sheet!";
				return strMessage;
			}
			
			// meals sheet
			XSSFSheet mealSheet = book.getSheetAt(1);
			XSSFRow mealRow = mealSheet.getRow(1);
			
			if (!mealRow.getCell(1).getStringCellValue().equals("No.")
					|| !mealRow.getCell(2).getStringCellValue().equals("WBS")
					|| !mealRow.getCell(3).getStringCellValue().equals("Amount")
					|| !mealRow.getCell(4).getStringCellValue().equals("On")
					|| !mealRow.getCell(5).getStringCellValue().equals("Reason")
					|| !mealRow.getCell(6).getStringCellValue().equals("Restaurant")
					|| !mealRow.getCell(7).getStringCellValue().equals("Number of Attendees")
					|| !mealRow.getCell(8).getStringCellValue().equals("Internal attendees")) {
				strMessage = "The selected file has wrong format int the meals sheet!";
				return strMessage;
			}
			
		} catch (Exception e) {
			strMessage = "The selected file has somethong wrong!";
		}
		
		logger.info("check result:" + strMessage);
		
		return strMessage;
		
	}
	
	/**
	 * get cell data
	 * 
	 * @param cell
	 * @return
	 */
	private static String getStringFromCell (XSSFCell cell) {
		String result = "";
		
		if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
			result = cell.getStringCellValue();
		} else if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
			try {
				double value = cell.getNumericCellValue();
				result = String.valueOf(value);
			} catch (Exception e) {
				result = "0";
			}
		} else {
			result = "";
		}
		
		return result;
	}
	
	/**
	 * set cell value. if cell not exists, create it
	 * 
	 * @param row
	 * @param iColumn
	 * @param strValue
	 */
	private static void setValueToCell(XSSFRow row, int iColumn, String strValue) {
		
		XSSFCell cell = row.getCell(iColumn);
		
		if (cell == null) {
			cell = row.createCell(iColumn);
		}
		
		cell.setCellValue(strValue);
		
	}

}
