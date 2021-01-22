import com.fasterxml.jackson.core.FormatSchema;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.*;
import java.util.function.Supplier;
import java.util.stream.Collectors;
import java.util.stream.Stream;

import static java.util.stream.Collectors.toMap;


public class ConvertExcelToJson {

    public static void main(String[] args) throws IOException {
        List<Map<String, String>> banks = readExcelFile("banks.xlsx");
        writeObjects2JsonFile(banks, "banks.json");
    }

    private static List<Map<String, String>> readExcelFile(String filePath) throws IOException {

        FileInputStream excelFile = new FileInputStream(new File(filePath));
        Workbook workbook = new XSSFWorkbook(excelFile);
        Sheet sheet = workbook.getSheetAt(0);
        Iterator<Row> rows = sheet.iterator();
        List<String> hederList = new ArrayList<>();
        List<Map<String, String>> result = new ArrayList<>();

//            Supplier<Stream<Row>> rowStreamSupplier = Util.getRowStreamSupplier(sheet);
//
//            Row headerRow = rowStreamSupplier.get().findFirst().get();
//            List<String> headerCells = Util.getStream(headerRow)
//                    .map(Cell::getStringCellValue)
//                    .collect(Collectors.toList());
//
//            int colCount = headerCells.size();
//
//            return rowStreamSupplier.get()
//                    .skip(1)
//                    .map(row -> {
//
//                        List<String> cellList = Util.getStream(row)
//                                .map(Cell::getStringCellValue)
//                                .collect(Collectors.toList());
//
//                        return Util.cellIteratorSupplier(colCount)
//                                .get()
//                                .collect(toMap(headerCells::get, cellList::get));
//                    })
//                    .collect(Collectors.toList());

        int rowNumber = 0;
        while (rows.hasNext()) {
            Row currentRow = rows.next();
            // skip header
            if (rowNumber == 0) {
                Iterator<Cell> header = currentRow.iterator();
                while (header.hasNext()) {
                    Cell cell = header.next();
                    String value = cell.getStringCellValue();
                    if (value.isBlank()) {
                        continue;
                    }
                    hederList.add(cell.getStringCellValue());
                }
                rowNumber++;
                continue;
            }

            Iterator<Cell> cellsInRow = currentRow.iterator();
            int cellIndex = 0;

            Map<String, String> map = new LinkedHashMap<>();

            while (cellIndex<hederList.size()) {
                Cell currentCell = cellsInRow.next();
                map.put(hederList.get(cellIndex), currentCell.getStringCellValue());
                cellIndex++;
            }
            result.add(map);
        }
        // Close WorkBook
        workbook.close();
        return result;

    }

    private static void writeObjects2JsonFile(List<Map<String, String>> list, String pathFile) {
        ObjectMapper mapper = new ObjectMapper();
        File file = new File(pathFile);
        try {
            // Serialize Java object info JSON file.
            mapper.writerWithDefaultPrettyPrinter().writeValue(file, list);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

}
