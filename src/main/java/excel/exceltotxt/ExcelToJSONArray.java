package excel.exceltotxt;

import java.io.IOException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.json.JSONArray;

public class ExcelToJSONArray {
	private JSONArray table;
	private FormulaEvaluator evaluator;

	public ExcelToJSONArray(Workbook workbook, Sheet sheet) throws IOException {
		this.table = new JSONArray();
		evaluator = workbook.getCreationHelper().createFormulaEvaluator();
		for (int i = 0; i < sheet.getLastRowNum(); i++) {
			Row row = sheet.getRow(i);
			if (row != null) {
				JSONArray auxJSONArray = new JSONArray();
				for (int j = 0; j < row.getLastCellNum(); j++) {
					Cell cell = row.getCell(j);
					if (cell.getCellType() == CellType.NUMERIC) auxJSONArray.put(String.valueOf(cell.getNumericCellValue()));
					else if(cell.getCellType() == CellType.STRING) auxJSONArray.put(cell.getStringCellValue());
					else if(cell.getCellType() == CellType.BOOLEAN) auxJSONArray.put(String.valueOf(cell.getBooleanCellValue()));
					else if(cell.getCellType() == CellType.FORMULA) {
						if(evaluator.evaluateFormulaCell(cell) ==  CellType.NUMERIC) auxJSONArray.put(String.valueOf(cell.getNumericCellValue()));
						else if(evaluator.evaluateFormulaCell(cell) ==  CellType.STRING) auxJSONArray.put(cell.getStringCellValue());
						else if(evaluator.evaluateFormulaCell(cell) ==  CellType.BOOLEAN) auxJSONArray.put(String.valueOf(cell.getBooleanCellValue()));
						else auxJSONArray.put("");
					}
					else auxJSONArray.put("");
				}
				this.table.put(auxJSONArray);
			}
		}
	}

	public JSONArray getJSONArray() {
		return this.table;
	}
}