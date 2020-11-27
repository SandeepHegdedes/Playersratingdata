package com.sandeep.Excelhand;

import java.io.IOException;
import java.lang.reflect.Array;
import java.util.Arrays;
import java.util.Collections;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * 
 * @author Sandeep Hegde
 *
 */
public class Main 
{
	public  static void main(String[]args) throws IOException {
		ExcelHandler excelHandler = new ExcelHandler();
		excelHandler.getRowCount();
		excelHandler.getRating();
		excelHandler.playersName();
		excelHandler.getMaxAge();
		excelHandler.getNameAndAge();
	
	
		
		System.out.println("Average ratings of players whose age is above 40 years is :");
		System.out.println(excelHandler.gettingAvgRating());
		System.out.println("The player name with the highest age to rating ratio");
		System.out.println(excelHandler.getHightAgeRatingRatio());
        System.out.println("names of 3 players with least ratings");
		System.out.print(Arrays.toString(excelHandler.getLeastRatingPlayers()));


}
}






