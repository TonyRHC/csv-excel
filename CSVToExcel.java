import java.io.File;
import java.io.FileOutputStream;
import java.nio.charset.Charset;
import java.util.Map;

import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVParser;
import org.apache.commons.csv.CSVRecord;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CSVToExcel {
	public void generateExcel (File csvFile, String outputName, String outputDirectory) {
		try {
			CSVParser parser = CSVParser.parse(csvFile, Charset.defaultCharset(), CSVFormat.DEFAULT.withFirstRecordAsHeader());
			FileOutputStream fileOut = new FileOutputStream(outputDirectory + outputName + ".xlsx");
			Workbook wb = new XSSFWorkbook();
			Sheet sheet = wb.createSheet();
			
			boolean firstTime = true;
			int rowCount = 1;
			for (CSVRecord csvRecord : parser) {
				if (firstTime) {
					Row headerRow = sheet.createRow(0);
					int headerColumnCount = 0;
					for (String header : csvRecord.toMap().keySet()) {
						Cell cell = headerRow.createCell(headerColumnCount);
						cell.setCellValue(header);
						headerColumnCount++;
					}
					firstTime = false;
				}
				
				Map<String, String> record = csvRecord.toMap();
				Row row = sheet.createRow(rowCount);
				
				int columnCount = 0;
				for (String key : record.keySet()) {
					Cell cell = row.createCell(columnCount);
					cell.setCellValue(record.get(key));
					columnCount++;
				}
				
				rowCount++;
			}
			
			wb.write(fileOut);
			wb.close();
			fileOut.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}
