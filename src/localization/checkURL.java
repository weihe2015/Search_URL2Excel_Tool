package localization;

/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.*;
/**
 *
 * @author wei7771
 */
public class checkURL {
    public static final String userDir = System.getProperty("user.dir");
    public static String[][] valueArray = null;
    public static String[][] desktopArray = null; 
    public static ArrayList<String> pubList = null;
    public static String lang = "";
    public static String HOName = "";
    public static String outputFileName = "";
    public static String outputFilePath = "";
    public static String desktop = "";
    public static String server = "";
    public static void check(String desktopFile, String serverFile, String inputFolder){
        
        try{
            String desktopFolder = desktopFile.substring(desktopFile.lastIndexOf("\\")+1, desktopFile.length());
            desktop = desktopFolder.substring(0, desktopFolder.indexOf("_"));
            String serverFolder = serverFile.substring(serverFile.lastIndexOf("\\")+1,serverFile.length());
            server = serverFolder.substring(0, serverFolder.indexOf("_"));
            
            pubList = new ArrayList<>();
            searchFile(inputFolder);
            String parFolder = inputFolder.substring(0,inputFolder.lastIndexOf("\\"));
            HOName = parFolder.substring(parFolder.lastIndexOf("\\")+1, parFolder.lastIndexOf("\\")+5);
            lang = inputFolder.substring(inputFolder.lastIndexOf("\\")+1,inputFolder.length());
            outputFileName = parFolder.substring(parFolder.lastIndexOf("\\")+1,parFolder.length());
            outputFilePath = parFolder + "\\" + outputFileName + "_"+ lang +".xlsx";
            
            valueArray = new String[pubList.size()+1][6];
            valueArray[0][0] = "Language";
            valueArray[0][1] = "HO#";
            valueArray[0][2] = "Publication Name";
            valueArray[0][3] = "Type";
            valueArray[0][4] = "Topic Name";
            valueArray[0][5] = "URL";
            for(int i = 0; i < pubList.size(); i++){        
                String fullPath = pubList.get(i);
                valueArray[i+1][0] = lang.toUpperCase().trim();
                valueArray[i+1][1] = HOName.trim();
                valueArray[i+1][4] = fullPath.substring(fullPath.lastIndexOf("\\")+1,fullPath.length());
                if(fullPath.contains("\\topic\\")){
                    valueArray[i+1][3] = "topic";
                    valueArray[i+1][2] = fullPath.substring(fullPath.indexOf("\\P")+4,fullPath.indexOf("\\topic\\")).trim();
                }
                else if(fullPath.contains("\\map\\")){
                    valueArray[i+1][3] = "map";
                    valueArray[i+1][2] = fullPath.substring(fullPath.indexOf("\\P")+4,fullPath.indexOf("\\map\\")).trim();
                }
            }
            
               /* for(int i = 0; i < valueArray.length; i++){
                    for(int j = 0; j < valueArray[i].length; j++){
                        System.out.print(valueArray[i][j] + ",");
                    }
                    System.out.println();
            }*/

            
            File inputDesktopFile = new File(desktopFile);
            File inputServerFile = new File(serverFile);
            XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream(inputDesktopFile));
            XSSFSheet mysheet = workbook.getSheetAt(0);
            int desktopRowNum = mysheet.getLastRowNum();
            
            XSSFWorkbook serverWorkbook = new XSSFWorkbook(new FileInputStream(inputServerFile));
            XSSFSheet serverSheet = serverWorkbook.getSheetAt(0);
            int serverRowNum = serverSheet.getLastRowNum();
            
            for(int k = 1; k < valueArray.length; k++){
                //System.out.println(valueArray[k][3]);
                if(valueArray[k][3].equals("topic") && (!valueArray[k][4].trim().startsWith("cfg"))){
                   //System.out.println(k + " " +valueArray[k][3]);
                   String source = valueArray[k][4].trim();
                   for(int i=1; i<desktopRowNum+1; i++){
                        Row desktopRow = mysheet.getRow(i);
                        if(desktopRow != null){
                           String targetString = desktopRow.getCell(2).getStringCellValue().trim();
                           if(source.contains(targetString)){
                               String desktopURL = desktopRow.getCell(0).getStringCellValue().trim();
                               desktopURL = desktopURL.replace(".com/en\\", ".com/"+lang.toLowerCase()+"\\");
                               desktopURL = desktopURL.replace("http://"+desktop,"http://"+desktop+"uat");
                               valueArray[k][5] = desktopURL;
                            }
                        }
                    }
                   
                   for(int j= 0; j<serverRowNum+1; j++){
                       Row serverRow = serverSheet.getRow(j);
                       if(serverRow != null){
                           String targetString1 = serverRow.getCell(2).getStringCellValue().trim();
                           if(source.contains(targetString1)){
                               String serverURL = serverRow.getCell(0).getStringCellValue().trim();
                               serverURL = serverURL.replace("/en\\", "/"+lang.toLowerCase()+"\\");
                               serverURL = serverURL.replace("http://"+server,"http://"+server+"uat");
                              // System.out.println(serverURL);
                               if(valueArray[k][5] != null){
                                   valueArray[k][5] = valueArray[k][5] + "\n" + serverURL;
                               }
                               else{
                                   valueArray[k][5] = serverURL;
                               }
                               
                           }
                       }
                   }
            }
            
            XSSFWorkbook outputworkbook = new XSSFWorkbook();
            XSSFSheet outputsheet = outputworkbook.createSheet("sheet1");
            XSSFCellStyle outputstyle = outputworkbook.createCellStyle();
            outputstyle.setWrapText(true);
            int outputRowNum = 0;
            int outputCellNum = 0;
            for(int i = 0; i<valueArray.length; i++){
                Row outputRow = outputsheet.createRow(outputRowNum++);
                for(int j=0; j < valueArray[1].length; j++){
                    Cell outputCell = outputRow.createCell(outputCellNum++);
                    if(valueArray[i][j] != null){
                        outputCell.setCellValue(valueArray[i][j]);
                    }
                    else{
                        outputCell.setCellValue("N/A");
                    }
                    if(j == 5){
                         //outputsheet.autoSizeColumn(4);
                         outputCell.setCellStyle(outputstyle);
                    }  
                }
                    outputCellNum = 0;
            }
            outputsheet.autoSizeColumn(2);
            outputsheet.autoSizeColumn(4);
            outputsheet.autoSizeColumn(5);
            FileOutputStream out = new FileOutputStream(new File(outputFilePath));
            outputworkbook.write(out);
            out.close();
            }
        }catch(Exception e){
            try{
                File file = new File(userDir+"\\log.txt");
                if(!file.exists()){
                    file.createNewFile();
                }
                FileWriter fw = new FileWriter(file.getAbsoluteFile());
                BufferedWriter bw = new BufferedWriter(fw);
                bw.write(e.getMessage());
                bw.write(e.getLocalizedMessage());
                bw.close();
                fw.close();
                }catch(Exception e1){
                    e1.printStackTrace();
                }     
            e.printStackTrace();
        }
    }
    
    public static void searchFile(String folder){
        File dir = new File(folder);
        File[] dirs = dir.listFiles();
        if(dirs != null){
            for(File child: dirs){
                String filePath = child.getAbsolutePath();
                if(filePath.endsWith(".zip")){
                    return;
                }
                else if(child.isDirectory()){
                    searchFile(filePath);
                }
                else if(filePath.endsWith(".xml")){
                    pubList.add(filePath);
                }
            }
        }
    }
       
  /*  public static void main(String args[]){
        String inputFolder = "C:\\Users\\wei7771\\Desktop\\HO10_web_desktop\\de";
        String desktopName = "C:\\Users\\wei7771\\Desktop\\10.3.1_URL\\desktop_urls_08172015.xlsx";
        String serverName = "C:\\Users\\wei7771\\Desktop\\10.3.1_URL\\server_urls_08172015.xlsx";
        check(desktopName,serverName,inputFolder);
    } */
}
