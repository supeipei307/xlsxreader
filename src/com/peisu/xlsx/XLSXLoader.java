package com.peisu.xlsx;

import java.io.File;
//import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
//import java.io.InputStream;
import java.text.SimpleDateFormat;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class XLSXLoader {

    //Define the variables
    //private static InputStream inputStream;
	private static OPCPackage opcPackage;
    private static XSSFWorkbook xssfWorkbook;
    private static FormulaEvaluator formulaEvaluator;
    private static int maxCellCount  = 0;
    private static int maxRowCount = 0;

    //args[0]: the path to the xlsx file
    //args[1]: the sheet number to process, start from 0
    //args[2]: the date format output to standard output, for example, yyyy-MM-dd
    //args[3]: start from which line to read, if there is a header, start from line 1
    //args[4]: end to which line, 0 means to the end of the xlsx
    public static void main(String[] args) {
        
        try {
            //Open the input stream from the file, initialize the workbook
            //inputStream = new FileInputStream(args[0]);
            //xssfWorkbook = new XSSFWorkbook(inputStream);  
        	opcPackage = OPCPackage.open(new File(args[0]));
        	xssfWorkbook = new XSSFWorkbook(opcPackage);
        	
            formulaEvaluator = xssfWorkbook.getCreationHelper().createFormulaEvaluator();
            
            //Open the sheet
            XSSFSheet xssfSheet = xssfWorkbook.getSheetAt(Integer.valueOf(args[1]));
            if (xssfSheet != null){
                //Get the first row of the sheet
                XSSFRow firstXSSFRow = xssfSheet.getRow(0);
                if (firstXSSFRow != null){
                    //Set the max cell count to be the last cell number of the first line
                    maxCellCount = firstXSSFRow.getLastCellNum();
                    
                    if (Integer.valueOf(args[4]) <= 0)
                        maxRowCount = xssfSheet.getLastRowNum();
                    else
                    	maxRowCount = Integer.valueOf(args[4]);
                    //Loop to read the rows
                    for (int rowNum = Integer.valueOf(args[3]);rowNum <= ((maxRowCount <= xssfSheet.getLastRowNum())?maxRowCount:xssfSheet.getLastRowNum());rowNum++) {
                        //Get the row
                        XSSFRow xssfRow = xssfSheet.getRow(rowNum);
                        
                        if (xssfRow != null){
                            //Loop to read the cells
                            for (int cellNum = 0;cellNum < maxCellCount;cellNum++){
                                //Get the cell
                                XSSFCell xssfCell = xssfRow.getCell(cellNum);
                                
                                if (xssfCell != null){
                                    //Process the cell based on the cell type
                                    switch (xssfCell.getCellTypeEnum()){
                                        case STRING:
                                            System.out.print(xssfCell.getStringCellValue());
                                            break;
                                        case NUMERIC:
                                            //If the cell matches the date format, output the cell as a date
                                            if (DateUtil.isCellDateFormatted(xssfCell)) {
                                                SimpleDateFormat dateFormat = new SimpleDateFormat(args[2]);
                                                System.out.print(dateFormat.format(xssfCell.getDateCellValue()));
                                            } 
                                            else
                                                System.out.print(xssfCell.getNumericCellValue());
                                            break;
                                        case BOOLEAN:
                                            System.out.print(xssfCell.getBooleanCellValue());
                                            break;
                                        case FORMULA:
                                            //For formula cell, evaluate the formula to get the result
                                            CellValue cellValue = formulaEvaluator.evaluate(xssfCell); 
                                            //Process the formula cell based on the type of the result
                                            switch(cellValue.getCellTypeEnum()){
                                                case STRING:
                                                    System.out.print(xssfCell.getStringCellValue());
                                                    break;
                                                case NUMERIC:
                                                    //If the result matches the date format, output the result as a date
                                                    if (DateUtil.isCellDateFormatted(xssfCell)) {
                                                        SimpleDateFormat dateFormat = new SimpleDateFormat(args[2]);
                                                        System.out.print(dateFormat.format(xssfCell.getDateCellValue()));
                                                    }
                                                    else
                                                        System.out.print(xssfCell.getNumericCellValue());
                                                    break;
                                                case BOOLEAN:
                                                    System.out.print(xssfCell.getBooleanCellValue());
                                                    break;
                                                default:
                                                    System.out.print(xssfCell.getRawValue());
                                            }
                                            break;
                                        case ERROR:
                                            //System.out.print(xssfCell.getErrorCellString());
                                            System.out.print("");
                                            break;
                                        default:
                                            System.out.print(xssfCell.getRawValue());
                                    }
                                }
                                //Add a column delimiter between the output cells
                                if(cellNum < maxCellCount - 1)
                                    System.out.print("\t");
                            }
                        }
                        //Add a row delimiter between the output rows
                        System.out.println("");
                    }
                }
            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } catch (InvalidFormatException e) {
			e.printStackTrace();
		} finally {
            try {
                //xssfWorkbook.close();
                opcPackage.close();
                //inputStream.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }
}
