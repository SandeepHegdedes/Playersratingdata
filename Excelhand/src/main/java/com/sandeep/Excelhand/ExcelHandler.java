package com.sandeep.Excelhand;

import java.io.IOException;
import java.lang.reflect.Array;
import java.util.Arrays;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * 
 * @author Sandeep
 *
 */
public class ExcelHandler {
  //Reading work book
	public  XSSFSheet getsheet() 
	{
		String excelpath=GlobalConstants.EXCEL_PATH;
		XSSFWorkbook workbook = null;
		try {
			workbook = new XSSFWorkbook(excelpath);
			workbook.close();
		} catch (IOException e) {
			e.printStackTrace();
		}
		XSSFSheet sheet= workbook.getSheet(GlobalConstants.SHEET_NAME);
		return sheet ;

	}
	
	//counting the number of rows
	public int getRowCount() throws IOException
	{  
		XSSFSheet sheet = getsheet();
		int rowCount =sheet.getPhysicalNumberOfRows();
		return rowCount ;

	}

	//Returns the array of rating
	public  int[]  getRating() throws IOException{
		int [] rating=new int [getRowCount()];
		for (int i=1;i<getRowCount();i++) {
			double value=((XSSFSheet) getsheet()).getRow(i).getCell(3).getNumericCellValue();
			rating[i]=(int) value;
		}
		return rating;
	}

	//Returns average for rating
	public  int gettingAvgRating() throws IOException{

		int [] rating = getRating();
		int count=0;
		int totalrating =0;
		for (int i=1;i<getRowCount();i++) {
			double age=((XSSFSheet) getsheet()).getRow(i).getCell(2).getNumericCellValue();
			if(age>40) {
				count=count+1;
				totalrating = totalrating + rating[i];
			}

		}
		return totalrating/count;
	}
	
	//Returns the array of players name
	public  String[] playersName() throws IOException{
		XSSFSheet sheet = getsheet();
		String [] allname = new String[getRowCount()];
		for (int i=1;i<getRowCount();i++) {
			String name=((XSSFSheet) getsheet()).getRow(i).getCell(1).getStringCellValue();
			allname [i]= name;
		}
		return allname ;
	}
	
    //Returns Array of age
	private  int[]  ageList() throws IOException{
		int [] allage = new int [getRowCount()];
		for (int i=1;i<getRowCount();i++) {
			int age=(int) ((XSSFSheet) getsheet()).getRow(i).getCell(2).getNumericCellValue();
			allage [i]= age; 	
		}
		return allage ;
	}
	// Returns Maximum age among players
	public  int getMaxAge() throws IOException{
		int [] allage = ageList();
		Arrays.sort(allage);
		int maxage = allage[getRowCount()-1];	
		return maxage;
	}
	
    // Get maximum age player
	public void getNameAndAge() throws IOException{
		int [] allage = ageList();
		int ind = 0;
		for (int i = 0; i < allage.length; i++) {
			if (allage[i] == getMaxAge()) {
				ind = i;
				break;
			}
		}

		String stringArr = Arrays.toString(allage); 
	}
	
	// Return three lowest rating players
	public int[] getThreeLowest(int[] array)
	{
		Arrays.sort(array);
		int[] lowestthree = new int[3];
		
		for (int i= 0; i < lowestthree.length; i++)
		{
			lowestthree[i] = array[i + 1];
		}
		
		return lowestthree;
	}
	
	//Returns index value
	public int getIndex(int[] array, int t)
	{
		int ind = 0;
		for (int i = 0; i < array.length; i++) {
			if (array[i] == t) {
				ind = i;
				break;
	}
		}
		
		return ind;
	}
	
	
	public int getIndexd(double [] array, double t)
	{
		int ind = 0;
		for (int i = 0; i < array.length; i++) {
			if (array[i] == t) {
				ind = i;
				break;
	}
		}
		
		return ind;
	}
	
	//Returns least rating players
    public String [] getLeastRatingPlayers() throws IOException {
			int [] ratinglist = getRating();		
			int[] lowesthtree = getThreeLowest(ratinglist);
			
			Arrays.sort(lowesthtree);
			
			String [] lowRatingNames =  new String[3];
			String [] names = playersName();
			
			for (int i = 0; i < lowRatingNames.length; i++)
			{
				lowRatingNames[i] = names[getIndex(ratinglist, lowesthtree[i])];
			}
			
			return lowRatingNames;
    }
    
    
    public double largest(double[] arr) 
    { 
        int i; 
        double max = arr[0];  
        for (i = 1; i < arr.length; i++) 
            if (arr[i] > max) 
                max = arr[i]; 
       
        return max; 
    } 
    // returns index of maximum age to rating ratio
    public int getratioind() throws IOException
    {
    	int[] age = ageList();
    	int[] rating = getRating();
    	
    	double[] temp = new double [getRowCount()];
    	
    	for (int i =1 ; i < age.length; i++)
    	{
    		temp[i] = (double) ((double)age[i]/ (double)rating[i]);
    	}
    	
    	double [] temp1 = temp;
    	int ind = getIndexd(temp, largest(temp));
    	return ind;
    	
    	
    }
    // Returns player name having maximum age to rating ratio
    public String getHightAgeRatingRatio() throws IOException
    {
    	String [] playersname = playersName();
    	return playersname[getratioind()];
    }
}

