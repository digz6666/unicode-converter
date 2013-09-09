package mn.digz.unicodeconverter.excel;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.nio.file.FileSystems;
import java.nio.file.Files;
import java.nio.file.Path;
import mn.digz.unicodeconverter.lang.MongolianLanguageUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.crypt.Decryptor;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.filesystem.NPOIFSFileSystem;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author MethoD
 */
public class ExcelConverter {
    
    public static void convert(String input, String output) throws Exception {
        if(input != null && output != null) {
            Path inputPath = FileSystems.getDefault().getPath(input);
            Path outputPath = FileSystems.getDefault().getPath(output);
            String contentType = Files.probeContentType(inputPath);
            
            Workbook workbook = null;
            
            if(contentType.equals("application/vnd.ms-excel")) { // excel 2003
                //NPOIFSFileSystem poifs = new NPOIFSFileSystem(new FileInputStream(inputPath.toFile()));
                //EncryptionInfo decryptionInfo = new EncryptionInfo(poifs);
                //Decryptor decryptor = Decryptor.getInstance(decryptionInfo);
                //decryptor.verifyPassword("myhouse");
                //workbook = new HSSFWorkbook(decryptor.getDataStream(poifs));
                workbook = new HSSFWorkbook(new FileInputStream(inputPath.toFile()));
            } else if(contentType.equals("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")) { // excel 2007+
                //POIFSFileSystem poifs = new POIFSFileSystem(new FileInputStream(inputPath.toFile()));
                //EncryptionInfo decryptionInfo = new EncryptionInfo(poifs);
                //Decryptor decryptor = Decryptor.getInstance(decryptionInfo);
                //decryptor.verifyPassword("myhouse");
                //workbook = new XSSFWorkbook(decryptor.getDataStream(poifs));
                workbook = new XSSFWorkbook(new FileInputStream(inputPath.toFile()));
            }
            
            if(workbook != null) {
                SXSSFWorkbook outputWorkbook = new SXSSFWorkbook();
                for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                    Sheet sheet = workbook.getSheetAt(i);
                    Sheet outputSheet = outputWorkbook.createSheet();
                    for (int j = 0; j <= sheet.getLastRowNum(); j++) {
                        Row row = sheet.getRow(j);
                        if(row != null) {
                            Row outputRow = outputSheet.createRow(j);
                            for (int k = 0; k <= row.getLastCellNum(); k++) {
                                Cell cell = row.getCell(k);
                                if(cell != null) {
                                    Cell outputCell = outputRow.createCell(k, cell.getCellType());
                                    if(cell.getCellType() == Cell.CELL_TYPE_STRING) {
                                        // convert to unicode and copy
                                        if(cell.getStringCellValue() != null) {
                                            outputCell.setCellValue(MongolianLanguageUtil.toUnicode(cell.getStringCellValue()));
                                        }
                                    } else {
                                        // just copy cell
                                        switch(cell.getCellType()) {
                                            case Cell.CELL_TYPE_BOOLEAN:
                                                outputCell.setCellValue(cell.getBooleanCellValue());
                                                break;
                                            case Cell.CELL_TYPE_FORMULA:
                                                outputCell.setCellValue(cell.getCellFormula());
                                                break;
                                            case Cell.CELL_TYPE_NUMERIC:
                                                outputCell.setCellValue(cell.getNumericCellValue());
                                                break;
                                            case Cell.CELL_TYPE_ERROR:
                                                outputCell.setCellValue(cell.getErrorCellValue());
                                                break;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                outputWorkbook.write(new FileOutputStream(outputPath.toFile()));
            }
        }
    }
}
