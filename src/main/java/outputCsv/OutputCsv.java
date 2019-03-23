package outputCsv;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStreamReader;
import java.io.PrintWriter;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.yaml.snakeyaml.Yaml;

public class OutputCsv {

	public static void main(String[] args) {


		// ファイル名取得
		Map<Map<String,String>,Map<String,String>> properties = getProperties("setting.yml");
		Map<String,String> path = properties.get("path");
		String inputXlxsPath = path.get("INPUT_PATH");
		String outputCsvPath = path.get("OUTPUT_PATH");
		Map<String,String>  sheets = properties.get("sheet");
		Map<String,String>  outputcolumn = properties.get("outputcolumn");


		// TODO 出力列名指定

		try {
			// Excel読み込み
			writeCsv(readXlxs(inputXlxsPath,sheets,outputcolumn),outputCsvPath);

		} catch (Exception e) {

			// TODO ログ出力を追加
			e.printStackTrace();

		}
	}

	private static List<String> readXlxs(String inputXlxsPath,Map<String,String> targetSheetNames,Map<String,String> targetColumns) throws EncryptedDocumentException, IOException, InvalidFormatException {

		List<String> resultList = new ArrayList<String>();

		// 対象Excel読み込み
		Workbook targetBook = WorkbookFactory.create(new File(inputXlxsPath));

		// 複数シート対応
		for(String key:targetSheetNames.keySet()) {
			Sheet targetSheet = targetBook.getSheet(targetSheetNames.get(key));

			int lastRowNumber = targetSheet.getLastRowNum();

			for (int i = 1; i <= lastRowNumber; i++) {
				Row targetRow = targetSheet.getRow(i);

				// 空行の場合は次の行へ
				if (targetRow == null) {
					continue;
				}

				String temp = "";
				for(String index:targetColumns.keySet()) {
					temp += targetRow.getCell(Integer.parseInt(targetColumns.get(index))).toString();
				}
				resultList.add(temp);

			}
		}

		return resultList;
	}

	/**
	 * データ出力(CSV)を行う
	 * @param
	 * @throws FileNotFoundException
	 */
	private static void writeCsv(List<String> outputList,String outputPath) throws FileNotFoundException {
		File file = new File(outputPath);

		PrintWriter pw = new PrintWriter(file);
		for (String element : outputList) {
			pw.println(element);
		}
		pw.close();
	}

	private static Map<Map<String,String>,Map<String,String>> getProperties(String key) {

		Map<Map<String,String>,Map<String,String>> result = new HashMap();

		InputStreamReader reader;

	    reader = new InputStreamReader(ClassLoader.getSystemResourceAsStream(key));

	    Yaml yaml = new Yaml();
	    result = (Map<Map<String, String>, Map<String, String>>) yaml.load(reader);

		return result;
	}

}

