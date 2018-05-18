package org.whissper;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.text.SimpleDateFormat;
import javax.naming.InitialContext;
import javax.naming.NamingException;
import javax.sql.DataSource;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * ExcelLodaerEngine class
 */
public class ExcelLoaderEngine {
    private String resultStr;
    
    private String path;
    private File xlsxFile;
    private String startDate;
    private String endDate;
    private String depID;
    
    private XSSFWorkbook workbook;
    private XSSFSheet sheet;
    
    private Row row;
    private Cell cell;
    private int rowNum = 0;
    
    private static final String MYSQL_DATA_SOURCE_STRING = "jdbc/exitlogMySQL";
    
    public ExcelLoaderEngine(String pathVal, String startDateVal, String endDateVal, String depIDVal){
        this.path = pathVal;
        this.startDate = startDateVal;
        this.endDate = endDateVal;
        this.depID = depIDVal;
        
        resultStr = "";
    }
    
    private String getDayName(int dayNum){
        String dayName = "Вс";
        switch(dayNum){
            case 1:
                dayName = "Вс";
                break;
            case 2:
                dayName = "Пн";
                break;
            case 3:
                dayName = "Вт";
                break;
            case 4:
                dayName = "Ср";
                break;
            case 5:
                dayName = "Чт";
                break;
            case 6:
                dayName = "Пт";
                break;    
            case 7:
                dayName = "Сб";
                break;  
        }
        return dayName;
    }
    
    private void alertError(String context){
        resultStr = "ERROR_JAVA|"+context;
    }
    
    private void initWorkbook(){
        workbook = new XSSFWorkbook();
        sheet = workbook.createSheet("Данные по выходам c "+ startDate +" по "+ endDate);
    }
    
    private void fillExcelHeader(){
        row = sheet.createRow(rowNum++);
                
        cell = row.createCell(0);
        cell.setCellValue("День");
        cell = row.createCell(1);
        cell.setCellValue("Дата");
        cell = row.createCell(2);
        cell.setCellValue("Сотрудник");
        cell = row.createCell(3);
        cell.setCellValue("Объекты");
        cell = row.createCell(4);
        cell.setCellValue("Примечание");
        cell = row.createCell(5);
        cell.setCellValue("Цель выхода");
        cell = row.createCell(6);
        cell.setCellValue("С");
        cell = row.createCell(7);
        cell.setCellValue("До");
        cell = row.createCell(8);
        cell.setCellValue("Часы");
    }
    
    private void fillExcelRows(){
        initWorkbook();
        fillExcelHeader();
        
        InitialContext ctx;
        DataSource ds;
        
        Connection con       = null;
        PreparedStatement st = null;
        ResultSet rs         = null;
        
        try {
            ctx = new InitialContext();
            ds = (DataSource)ctx.lookup(MYSQL_DATA_SOURCE_STRING);
        } catch (NamingException ex) {
            alertError("Exception -- fillExcelRows() -- : " + ex);
            return;
        }
        
        try {
            con = ds.getConnection();
            st  = con.prepareStatement( 
                "SELECT "+
                    "`exits`.`id`, "+
                    "DAYOFWEEK(`exits`.`date`) AS 'dayofweek', "+
                    "`exits`.`date`, "+ 
                    "`users`.`fio`, "+ 
                    "`objects`.`name` AS 'objectname', "+
                    "`objects`.`note` AS 'objectnote', "+
                    "`objects`.`postal_index` AS 'objectpostalindex', "+
                    "`objects`.`region` AS 'objectregion', "+
                    "`objects`.`town` AS 'objecttown', "+
                    "`objects`.`street` AS 'objectstreet', "+
                    "`objects`.`building` AS 'objectbuilding', "+
                    "`objects`.`apartment` AS 'objectapartment', "+
                    "`objects`.`geo_lat` AS 'objectgeolat', "+
                    "`objects`.`geo_lon` AS 'objectgeolon', "+
                    "`objects`.`old_format` AS 'objectoldformat', "+
                    "`points`.`name` AS 'point', "+ 
                    "`exits`.`point_description`, "+
                    "`exits`.`time_exit`, "+ 
                    "`exits`.`time_return`, "+ 
                    "TIMEDIFF(`exits`.`time_return`, `exits`.`time_exit`) AS 'hours' "+
                "FROM `exits` "+
                "LEFT JOIN `users` ON `users`.`id`=`exits`.`user_id` "+
                "LEFT JOIN `objects` ON `objects`.`exit_id` = `exits`.`id` "+
                "LEFT JOIN `points` ON `points`.`id` = `exits`.`point_id` "+ 
                "WHERE "+ 
                    "`exits`.`deleted` = 0 AND "+ 
                    "`exits`.`date` >= ? AND "+
                    "`exits`.`date` <= ? AND "+
                    "`users`.`department_id` = ? "+
                "ORDER BY `exits`.`id` DESC" );
            st.setString(1, startDate);
            st.setString(2, endDate);
            st.setString(3, depID);

            rs = st.executeQuery();
            
            int currentId = -1;
            int currentPosUniqueRow = -1;
            
            while(rs.next()){
                row = sheet.createRow(rowNum);
                
                if(currentId != rs.getInt("id")){
                    if(currentPosUniqueRow != -1 && ((rowNum-1)-currentPosUniqueRow) > 0){
                        sheet.addMergedRegion(new CellRangeAddress(currentPosUniqueRow, rowNum-1, 0, 0));
                        sheet.addMergedRegion(new CellRangeAddress(currentPosUniqueRow, rowNum-1, 1, 1));
                        sheet.addMergedRegion(new CellRangeAddress(currentPosUniqueRow, rowNum-1, 2, 2));
                        sheet.addMergedRegion(new CellRangeAddress(currentPosUniqueRow, rowNum-1, 5, 5));
                        sheet.addMergedRegion(new CellRangeAddress(currentPosUniqueRow, rowNum-1, 6, 6));
                        sheet.addMergedRegion(new CellRangeAddress(currentPosUniqueRow, rowNum-1, 7, 7));
                        sheet.addMergedRegion(new CellRangeAddress(currentPosUniqueRow, rowNum-1, 8, 8));
                    }
                    cell = row.createCell(0);
                    cell.setCellValue(getDayName(rs.getInt("dayofweek")));
                    cell = row.createCell(1);
                    cell.setCellValue(new SimpleDateFormat("dd.MM.yyyy").format(rs.getDate("date")));
                    cell = row.createCell(2);
                    cell.setCellValue(rs.getString("fio"));
                    cell = row.createCell(3);
                    if ( rs.getInt("objectoldformat") == 1 ) {
                        cell.setCellValue(rs.getString("objectname"));
                    } else {
                        cell.setCellValue(     rs.getString("objectstreet") +
                                           ", "+rs.getString("objectbuilding") +
                                           (rs.getString("objectapartment")==null ? "" : (", "+rs.getString("objectapartment"))) );
                    }
                    cell = row.createCell(4);
                    cell.setCellValue(rs.getString("objectnote"));
                    cell = row.createCell(5);
                    cell.setCellValue(rs.getString("point") +" "+ (rs.getString("point_description")==null ? "" : rs.getString("point_description")));
                    cell = row.createCell(6);
                    cell.setCellValue(new SimpleDateFormat("HH:mm").format(rs.getTime("time_exit")));
                    cell = row.createCell(7);
                    cell.setCellValue(new SimpleDateFormat("HH:mm").format(rs.getTime("time_return")));
                    cell = row.createCell(8);
                    cell.setCellValue(new SimpleDateFormat("HH:mm").format(rs.getTime("hours")));
                    
                    currentId = rs.getInt("id");
                    currentPosUniqueRow = rowNum;
                } else {
                    cell = row.createCell(0);
                    cell = row.createCell(1);
                    cell = row.createCell(2);
                    cell = row.createCell(3);
                    if ( rs.getInt("objectoldformat") == 1 ) {
                        cell.setCellValue(rs.getString("objectname"));
                    } else {
                        cell.setCellValue(     rs.getString("objectstreet") +
                                           ", "+rs.getString("objectbuilding") +
                                           (rs.getString("objectapartment")==null ? "" : (", "+rs.getString("objectapartment"))) );
                    }
                    cell = row.createCell(4);
                    cell.setCellValue(rs.getString("objectnote"));
                    cell = row.createCell(5);
                    cell = row.createCell(6);
                    cell = row.createCell(7);
                    cell = row.createCell(8);
                }
                
                rowNum++;
            }
        } catch (SQLException ex) {
            alertError("Exception -- fillExcelRows() -- : " + ex);
            return;
        } finally {
            //close current ResultSet
            try{if(rs!=null)rs.close();}catch(SQLException ex){}
            //close current Statement
            try{if(st!=null)st.close();}catch(SQLException ex){}
            //close current Connection
            try{if(con!=null)con.close();}catch(SQLException ex){}
        }
        
        System.out.println();
    }
    
    private void decorateTable(){
        sheet.autoSizeColumn(0);
        sheet.autoSizeColumn(1);
        sheet.autoSizeColumn(2);
        sheet.autoSizeColumn(3);
        sheet.autoSizeColumn(4);
        sheet.autoSizeColumn(5);
        sheet.autoSizeColumn(6);
        sheet.autoSizeColumn(7);
        sheet.autoSizeColumn(8);
    }
    
    private void createFile(){
        xlsxFile = new File(this.path + "Данные_по_выходам_"+ startDate +"_"+ endDate +".xlsx");
        if(xlsxFile.getParentFile()!=null){ xlsxFile.getParentFile().mkdirs(); }// Will create parent directories if not exists
        try {
            xlsxFile.createNewFile();
        } catch (IOException ex) {
            alertError("Exception -- createFile() -- : " + ex);
        }
    }
    
    private void writeFile(){
        try {
            FileOutputStream outputStream = new FileOutputStream(xlsxFile);
            workbook.write(outputStream);
            workbook.close();
            resultStr = "getfile/" + xlsxFile.getName();
        } catch (FileNotFoundException ex) {
            alertError("Exception -- writeFile() -- : " + ex);
        } catch (IOException ex) {
            alertError("Exception -- writeFile() -- : " + ex);
        }
    }
    
    public String loadData(){
        fillExcelRows();
        if( resultStr.contains("ERROR_JAVA") ){
            return resultStr;
        }
        decorateTable();
        if( resultStr.contains("ERROR_JAVA") ){
            return resultStr;
        }
        createFile();
        if( resultStr.contains("ERROR_JAVA") ){
            return resultStr;
        }
        writeFile();
        
        return resultStr;
    }
    
}
