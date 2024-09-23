package org.example;

import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.nio.file.DirectoryStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.text.SimpleDateFormat;
import java.util.*;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONArray;
import org.json.JSONObject;

public class ExcelToJson {
    public static void main(String[] args) {
        String folderPath = "/home/kike/ExcelToJson/0Input";  // Cambia esto a la ruta de tu carpeta
        String outputFolderPath = "/home/kike/ExcelToJson/1Output";  // Cambia esto a la ruta de tu carpeta de salida
        processAllFilesInFolder(folderPath, outputFolderPath);
    }

    private static void processAllFilesInFolder(String folderPath, String outputFolderPath) {
        try (DirectoryStream<Path> stream = Files.newDirectoryStream(Paths.get(folderPath), "*.xlsx")) {
            for (Path entry : stream) {
                List<Map<String, String>> data = readExcel(entry.toString());
                JSONArray jsonArray = createJsonArray(data);
                String outputFilePath = generateOutputFilePath(entry, outputFolderPath);
                saveJsonToFile(jsonArray, outputFilePath);
                System.out.println("JSON guardado en: " + outputFilePath);
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static List<Map<String, String>> readExcel(String filePath) {
        List<Map<String, String>> data = new ArrayList<>();
        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = new XSSFWorkbook(fis)) {
            Sheet sheet = workbook.getSheetAt(0);
            Iterator<Row> rowIterator = sheet.iterator();

            Row headerRow = rowIterator.next();
            List<String> headers = new ArrayList<>();
            headerRow.forEach(cell -> headers.add(cell.getStringCellValue()));

            while (rowIterator.hasNext()) {
                Row currentRow = rowIterator.next();
                Map<String, String> map = new HashMap<>();
                for (int i = 0; i < headers.size(); i++) {
                    Cell cell = currentRow.getCell(i);
                    String cellValue = cell != null ? cell.toString() : "";
                    map.put(headers.get(i), cellValue);
                }
                data.add(map);
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
        return data;
    }

    private static JSONArray createJsonArray(List<Map<String, String>> data) {
        JSONArray jsonArray = new JSONArray();
        for (Map<String, String> record : data) {
            JSONObject jsonObject = new JSONObject();
            jsonObject.put("code", record.get("PBK - Strukturgruppe"));

            JSONObject attributes = new JSONObject();
            JSONArray relevantArray = new JSONArray();
            for (Map.Entry<String, String> entry : record.entrySet()) {
                if (!entry.getKey().equals("PBK - Strukturgruppe") && entry.getValue() != null && !entry.getValue().isEmpty()) {
                    relevantArray.put(entry.getValue());
                }
            }
            attributes.put("relevant", relevantArray);
            jsonObject.put("attributes", attributes);

            jsonArray.put(jsonObject);
        }
        return jsonArray;
    }

    private static void saveJsonToFile(JSONArray jsonArray, String filePath) {
        try (FileWriter file = new FileWriter(filePath)) {
            file.write(jsonArray.toString(2));
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static String generateOutputFilePath(Path inputFile, String outputFolderPath) {
        String fileName = inputFile.getFileName().toString();
        String baseName = fileName.substring(0, fileName.lastIndexOf('.'));
        String timestamp = new SimpleDateFormat("yyyyMMdd_HHmmss").format(new Date());
        return Paths.get(outputFolderPath, baseName + "_" + timestamp + ".txt").toString();
    }
}