package excel.exceltotxt;

import java.io.IOException;
import java.util.ArrayList;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.json.JSONArray;
import org.json.JSONObject;

public class TemplateValidator {
	private FormulaEvaluator evaluator;
	private JSONArray template;
	private ArrayList<Integer[]> errors;
	
	/*Example template:
	{
	   "template":[
	      {
	         "type":"NUMERIC",
	         "length":4
	      },
	      {
	         "type":"STRING",
	         "length":12
	      },
	      {
	         "type":"BOOLEAN",
	         "length":1
	      },
	      {
	         "type":"FORMULA",
	         "length":50
	      },
	      {
	         "type":"MIX",
	         "length":8
	      }
	   ]
	}
	*/
	
	/*Error codes:
	 * 1	Excel has more columns than the template
	 * 2	Type mismatch between cell and template
	 * 3	Cell is longer than template permits
	 */

	public TemplateValidator(Workbook workbook, Sheet sheet, JSONObject templateObject) throws IOException {
		evaluator = workbook.getCreationHelper().createFormulaEvaluator();
		template = templateObject.getJSONArray("template");
		errors = new ArrayList<Integer[]>();
		
		for (int i = 0; i < sheet.getLastRowNum(); i++) {
			Row row = sheet.getRow(i);
			if (row != null) {
				for (int j = 0; j < row.getLastCellNum(); j++) {
					Cell cell = row.getCell(j);
					
					if(template.length()<j) errors.add(new Integer[] {1,i,j});
					else {
						if (template.getJSONObject(j).getString("type") == "NUMERIC") {
							if (cell.getCellType() != CellType.NUMERIC || (cell.getCellType() == CellType.FORMULA && evaluator.evaluateFormulaCell(cell) !=  CellType.NUMERIC))
								errors.add(new Integer[] {2,i,j});
						}
						else if (template.getJSONObject(j).getString("type") == "STRING") {
							if (cell.getCellType() != CellType.STRING || (cell.getCellType() == CellType.FORMULA && evaluator.evaluateFormulaCell(cell) !=  CellType.STRING))
								errors.add(new Integer[] {2,i,j});
						}
						else if (template.getJSONObject(j).getString("type") == "BOOLEAN") {
							if (cell.getCellType() != CellType.BOOLEAN || (cell.getCellType() == CellType.FORMULA && evaluator.evaluateFormulaCell(cell) !=  CellType.BOOLEAN))
								errors.add(new Integer[] {2,i,j});
						}
						
						
						if (cell.getCellType() == CellType.NUMERIC || (cell.getCellType() == CellType.FORMULA && evaluator.evaluateFormulaCell(cell) ==  CellType.NUMERIC)) {
							if(template.getJSONObject(j).getInt("lenght") < String.valueOf(cell.getNumericCellValue()).length()) errors.add(new Integer[] {3,i,j});
						}
						else if (cell.getCellType() == CellType.STRING || (cell.getCellType() == CellType.FORMULA && evaluator.evaluateFormulaCell(cell) ==  CellType.STRING)) {
							if(template.getJSONObject(j).getInt("lenght") < cell.getStringCellValue().length()) errors.add(new Integer[] {3,i,j});
						}
						else if (cell.getCellType() == CellType.BOOLEAN || (cell.getCellType() == CellType.FORMULA && evaluator.evaluateFormulaCell(cell) ==  CellType.BOOLEAN)) {
							if(template.getJSONObject(j).getInt("lenght") < String.valueOf(cell.getBooleanCellValue()).length()) errors.add(new Integer[] {3,i,j});
						}
					}
				}
			}
		}		
	}
}
