package util;

import javafx.application.Platform;
import javafx.scene.control.Label;
import javafx.scene.control.ProgressBar;
import javafx.scene.control.TextArea;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.nio.file.NoSuchFileException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.HashMap;
import java.util.Map;

public class ExcelManager {
    private TextArea logArea;
    private ProgressBar progressBar;
    private Label progressText;

    public ExcelManager(TextArea txtAreaLog, ProgressBar progressbar, Label progressText) {
        this.logArea = txtAreaLog;
        this.progressBar = progressbar;
        this.progressText = progressText;

        progressbar.setProgress(0.0);
    }

    public void read(String filePath, RowReader rowProcessor) throws Exception {

        printLog("엑셀 파일을 읽는 중 입니다..." + filePath);
        printLog("\t시트를 읽습니다.");

        //파일을 읽기위해 엑셀파일을 가져온다
        FileInputStream fis = new FileInputStream(filePath);
        Workbook workbook = getExtension(filePath).equals("xls") ? new HSSFWorkbook(fis) : new XSSFWorkbook(fis);

        int rowindex = 0;
        int columnindex = 0;

        //시트 수 (첫번째에만 존재하므로 0을 준다)
        //만약 각 시트를 읽기위해서는 FOR문을 한번더 돌려준다
        Sheet sheet = workbook.getSheetAt(0);
        //행의 수
        int rows = sheet.getPhysicalNumberOfRows();
        for (rowindex = 1; rowindex < rows; rowindex++) {
            //행을 읽는다
            Row row = sheet.getRow(rowindex);

            printLog("\t\t" + (rows - 1) + " 중 " + rowindex + "번 행을 읽고 있습니다.");
            rowProcessor.processing(row, sheet);
        }

        printLog("엑셀 파일을 성공적으로 읽었습니다...");
    }

    public void write(String savePath, RowWriter rowWriter) throws IOException {
        Workbook workbook = getExtension(savePath).equals("xls") ? new HSSFWorkbook() : new XSSFWorkbook();

        Sheet sheet = workbook.createSheet();

        rowWriter.processing(sheet);

        FileOutputStream outputStream = new FileOutputStream(savePath);

        workbook.write(outputStream);
        outputStream.close();
    }

    public Map<String, String> getCodeMapCodeAsValue(String codeFilePath) throws Exception {
        Map<String, String> resultCodeMap = new HashMap<>();

        read(codeFilePath, (row, sheet) -> {
            int columnindex = 0;
            if (row != null) {
                int cells = row.getPhysicalNumberOfCells();
                //셀의 수
                if (cells >= 2) {
                    resultCodeMap.put(extractStringValue(row.getCell(1)), extractStringValue(row.getCell(0)));
                }
            }
        });

        return resultCodeMap;
    }

    public Map<String, String> getCodeMapCodeAsKey(String codeFilePath) throws Exception {
        Map<String, String> resultCodeMap = new HashMap<>();

        printLog("재고원본으로부터 재고 정보를 조회 중입니다...");
        read(codeFilePath, (row, sheet) -> {
            int columnindex = 0;
            if (row != null) {
                //셀의 수
                int cells = row.getPhysicalNumberOfCells();
                if (cells >= 2) {
                    Cell codeCell = row.getCell(0);
                    Cell quantityCell = row.getCell(7);

                    if (codeCell.getCellTypeEnum() == CellType.BLANK ||
                            codeCell.getCellTypeEnum() == CellType.ERROR) return;

                    resultCodeMap.put(extractStringValue(row.getCell(0)), extractStringValue(row.getCell(7)));
                }
            }

        });
        printLog("엑셀 파일을 성공적으로 읽었습니다...");

        return resultCodeMap;
    }

    public void writeCodeFile(Map<String, String> codeMap, String stockPath, String savePath) throws Exception {

        printLog("엑셀 파일을 읽는 중 입니다..." + savePath);

        //파일을 읽기위해 엑셀파일을 가져온다
        FileInputStream fis = new FileInputStream(stockPath);
        Workbook workbook = getExtension(stockPath).equals("xls") ? new HSSFWorkbook(fis) : new XSSFWorkbook(fis);

        CellStyle styleOfColor = workbook.createCellStyle();

        // 정렬
        styleOfColor.setAlignment(HorizontalAlignment.CENTER); //가운데 정렬
        // 배경색
        styleOfColor.setFillForegroundColor(IndexedColors.PALE_BLUE.getIndex());
        styleOfColor.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        int rowindex = 0;
        int columnindex = 0;

        printLog("\t시트를 읽습니다.");

        //시트 수 (첫번째에만 존재하므로 0을 준다)
        //만약 각 시트를 읽기위해서는 FOR문을 한번더 돌려준다
        Sheet sheet = workbook.getSheetAt(0);
        //행의 수
        int rows = sheet.getPhysicalNumberOfRows();
        for (rowindex = 1; rowindex < rows; rowindex++) {
//            printLog("\t\t" + (rows - 1) + " 중 " + rowindex + "번 행을 읽고 있습니다.");
            updateProgress((double) (rowindex + 1) / rows);
            //행을 읽는다
            Row row = sheet.getRow(rowindex);

            Cell cellProductCode = row.getCell(1, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
            Cell cellProductName = row.getCell(2);

            String key = cellProductName.getStringCellValue();
            String value = codeMap.get(key);

            if (value != null) {
                cellProductCode.setCellValue(value);
                cellProductCode.setCellStyle(styleOfColor);
                printLog("\t상품 : " + key + ", 코드 : " + value + "로 변경되었습니다.");
            }
        }

        try {
            FileOutputStream fileOutputStream = new FileOutputStream(savePath);
            workbook.write(fileOutputStream);
            fileOutputStream.close();
            printLog("\t" + savePath + " 작성했습니다...");
        } catch (FileNotFoundException e) {
            printLog(e.getLocalizedMessage());
            printLog("엑셀 쓰기에 실패했습니다...");
            return;
        }

        printLog("엑셀 파일을 성공적으로 읽었습니다...");
    }

    public String getExtension(String filePath) throws NoSuchFileException {
        // 파일 확장자 확인
        if (filePath.lastIndexOf('.') <= 0) {
            printLog("** 잘못된 파일입니다. **");
            throw new NoSuchFileException("확장자가 없습니다.");
        }

        return filePath;
    }

    public void printLog(String log) {
        if (logArea != null) {
            Platform.runLater(() -> {
                logArea.appendText(log + "\n");
            });
        }
    }

    public void updateProgress(Double progress) {
        if (progressBar != null) {
            Platform.runLater(() -> {
                progressBar.setProgress(progress);
                progressText.setText("진행도 - " + String.format("%.2f", progress * 100) + "%");
            });
        }
    }

    public void writeFullFiles(Map<String, String> codeToQuantity, String target, String absolutePath) {

        printLog("엑셀 파일을 읽는 중 입니다..." + target);

        //타겟 파일(쇼핑몰재고)을 읽기위해 엑셀파일을 가져온다
        Workbook targetWorkbook;

        try {
            FileInputStream fis = new FileInputStream(target);
            targetWorkbook = getExtension(target).equals("xls") ? new HSSFWorkbook(fis) : new XSSFWorkbook(fis);
        } catch (Exception e) {
            printLog(e.getLocalizedMessage());
            return;
        }

        Workbook newWorkbook = new XSSFWorkbook();
        Sheet newSheet = newWorkbook.createSheet();

        // 미등록상품변경분 엑셀
        Workbook unregisteredWorkbook = new XSSFWorkbook();
        Sheet unregSheet = unregisteredWorkbook.createSheet();

        int rowindex = 0;

        //시트 수 (첫번째에만 존재하므로 0을 준다)
        Sheet sheet = targetWorkbook.getSheetAt(0);

        // 미등록시트 헤더행 생성
        copyRow(sheet, sheet.getRow(0),
                newWorkbook, newSheet.createRow(newSheet.getPhysicalNumberOfRows()));
        copyRow(sheet, sheet.getRow(0),
                unregisteredWorkbook, unregSheet.createRow(unregSheet.getPhysicalNumberOfRows()));

        //행의 수
        int rows = sheet.getPhysicalNumberOfRows();
        for (rowindex = 1; rowindex < rows; rowindex++) {
//            printLog("\t\t" + (rows - 1) + " 중 " + rowindex + "번 행을 읽고 있습니다.");
            updateProgress((double) (rowindex + 1) / rows);
            //행을 읽는다
            Row row = sheet.getRow(rowindex);

            Cell cellProductCode = row.getCell(1, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
            Cell cellQuantity = row.getCell(2);

            String key =  extractStringValue(cellProductCode);
            String value = codeToQuantity.get(key);

            if (value != null) {
                copyRow(sheet, row,
                        newWorkbook, newSheet.createRow(newSheet.getPhysicalNumberOfRows()));
                cellQuantity.setCellValue(value);
                printLog("\t상품 : " + key + ", 재고 : " + value + "로 변경되었습니다.");
            } else {
                copyRow(sheet, row,
                        unregisteredWorkbook, unregSheet.createRow(unregSheet.getPhysicalNumberOfRows()));
            }
        }

        Calendar calendar = Calendar.getInstance();
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyyMMdd_hhmmss_");
        String prefix = simpleDateFormat.format(calendar.getTime());
        String savePath;

        try {
            savePath = absolutePath + File.separator + prefix + "반영재고파일.xlsx";
            FileOutputStream fileOutputStream = new FileOutputStream(savePath);
            newWorkbook.write(fileOutputStream);
            fileOutputStream.close();
            printLog("\t" + savePath + " 작성했습니다...");
        } catch (Exception e) {
            printLog(e.getLocalizedMessage());
            printLog("\t반영재고파일 엑셀 쓰기에 실패했습니다...");
        }

        try {
            savePath = absolutePath + File.separator + prefix + "미반영재고파일.xlsx";
            FileOutputStream fileOutputStream = new FileOutputStream(savePath);
            unregisteredWorkbook.write(fileOutputStream);
            fileOutputStream.close();
            printLog("\t" + savePath + " 작성했습니다...");
        } catch (Exception e) {
            printLog(e.getLocalizedMessage());
            printLog("\t미반영재고파일 엑셀 쓰기에 실패했습니다...");
        }

        printLog("재고 조사가 완료 되었습니다...");
    }

    // 시트 행 복사
    // Source from : https://stackoverflow.com/questions/5785724/how-to-insert-a-row-between-two-rows-in-an-existing-excel-with-hssf-apache-poi
    private void copyRow(Sheet sheet, Row sourceRow, Workbook unregisteredWorkbook, Row newRow) {

        // Loop through source columns to add to new row
        for (int i = 0; i < sourceRow.getLastCellNum(); i++) {
            // Grab a copy of the old/new cell
            Cell oldCell = sourceRow.getCell(i);
            Cell newCell = newRow.createCell(i);

            // If the old cell is null jump to next cell
            if (oldCell == null) {
                newCell = null;
                continue;
            }

            // Copy style from old cell and apply to new cell
            CellStyle newCellStyle = unregisteredWorkbook.createCellStyle();
            newCellStyle.cloneStyleFrom(oldCell.getCellStyle());
            newCell.setCellStyle(newCellStyle);

            // If there is a cell comment, copy
            if (oldCell.getCellComment() != null) {
                newCell.setCellComment(oldCell.getCellComment());
            }

            // If there is a cell hyperlink, copy
            if (oldCell.getHyperlink() != null) {
                newCell.setHyperlink(oldCell.getHyperlink());
            }

            // Set the cell data type
            newCell.setCellType(oldCell.getCellTypeEnum());

            // Set the cell data value
            switch (oldCell.getCellTypeEnum()) {
                case BLANK:
                    newCell.setCellValue(oldCell.getStringCellValue());
                    break;
                case BOOLEAN:
                    newCell.setCellValue(oldCell.getBooleanCellValue());
                    break;
                case ERROR:
                    newCell.setCellErrorValue(oldCell.getErrorCellValue());
                    break;
                case FORMULA:
                    newCell.setCellFormula(oldCell.getCellFormula());
                    break;
                case NUMERIC:
                    newCell.setCellValue(oldCell.getNumericCellValue());
                    break;
                case STRING:
                    newCell.setCellValue(oldCell.getRichStringCellValue());
                    break;
            }
        }

    }

    private String extractStringValue(Cell cell) {
        String value = null;

        switch (cell.getCellTypeEnum()) {
            case FORMULA:
                value = cell.getCellFormula();
                break;
            case NUMERIC:
                value = cell.getNumericCellValue() + "";
                break;
            case STRING:
                value = cell.getStringCellValue() + "";
                break;
            case BLANK:
                value = cell.getBooleanCellValue() + "";
                break;
            case ERROR:
                value = cell.getErrorCellValue() + "";
                break;
        }

        return value;
    }
}
