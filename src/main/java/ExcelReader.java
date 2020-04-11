import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

import java.util.ArrayList;
import java.util.List;

public class ExcelReader {
    public static final String INPUT_EXCEL_PATH = "./input.xlsx";
    public static final String OUTPUT_EXCEL_PATH = "./output.xlsx";
    public static final String[] columns = {"STT", "Câu hỏi", "Đáp Án A", "Đáp Án B", "Đáp Án C", "Đáp Án D", "Đáp Án Đúng"};

    public static List<Question> readExcel(int sheetIndex) throws Exception {
        List<Question> listQuestion = new ArrayList();
        Workbook workbook = WorkbookFactory.create(new File(INPUT_EXCEL_PATH));
        Sheet sheet = workbook.getSheetAt(sheetIndex);
        System.out.println("Working on sheet: " + sheet.getSheetName());
        DataFormatter dataFormatter = new DataFormatter();

        for (Row row : sheet) {
            if (row.getRowNum() < 4) {
                continue;
            }
            Question question = new Question();
            for (Cell cell : row) {
                String cellValue = dataFormatter.formatCellValue(cell);
                switch (cell.getColumnIndex()) {
                    case 0:
                        question.setSequence(cellValue);
                        break;
                    case 1:
                        question.setDescription(cellValue);
                        break;
                    case 2:
                        question.setAnswerA(cellValue);
                        break;
                    case 3:
                        question.setAnswerB(cellValue);
                        break;
                    case 4:
                        question.setAnswerC(cellValue);
                        break;
                    case 5:
                        question.setAnswerD(cellValue);
                        break;
                    case 6:
                        question.setCorrectAnswer(cellValue);
                        break;
                    default:
                        System.out.println("No data at " + cell.getColumnIndex());
                }
            }
            listQuestion.add(question);
        }
        return listQuestion;
    }


    public static List<Data> readExcel() throws Exception {
        List<Data> masterList = new ArrayList<>();
        Workbook workbook = WorkbookFactory.create(new File(INPUT_EXCEL_PATH));

        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            Sheet sheet = workbook.getSheetAt(i);
            System.out.println("Working on sheet: " + sheet.getSheetName());
            DataFormatter dataFormatter = new DataFormatter();
            List<Question> listQuestion = new ArrayList();

            for (Row row : sheet) {
                if (row.getRowNum() < 4) {
                    continue;
                }
                Question question = new Question();
                for (Cell cell : row) {
                    String cellValue = dataFormatter.formatCellValue(cell);
                    switch (cell.getColumnIndex()) {
                        case 0:
                            question.setSequence(cellValue);
                            break;
                        case 1:
                            question.setDescription(cellValue);
                            break;
                        case 2:
                            question.setAnswerA(cellValue);
                            break;
                        case 3:
                            question.setAnswerB(cellValue);
                            break;
                        case 4:
                            question.setAnswerC(cellValue);
                            break;
                        case 5:
                            question.setAnswerD(cellValue);
                            break;
                        case 6:
                            question.setCorrectAnswer(cellValue);
                            break;
                        default:
                            System.out.println("No data at " + cell.getColumnIndex());
                    }
                }
                listQuestion.add(question);
            }
            Data lstQuestion = new Data(sheet.getSheetName(), listQuestion);
            masterList.add(lstQuestion);
        }
        return masterList;
    }


    public static void writeExcel(List<Question> listQuestion, String sheetName) throws Exception {
        Workbook workbook;
        try {
            workbook = new XSSFWorkbook(new FileInputStream(OUTPUT_EXCEL_PATH));
        } catch (Exception ex) {
            workbook = new XSSFWorkbook();
        }

        Sheet sheet = workbook.createSheet(sheetName);

        Row headerRow = sheet.createRow(0);
        for (int i = 0; i < columns.length; i++) {
            headerRow.createCell(i).setCellValue(columns[i]);
        }

        int rowNum = 1;
        for (Question question : listQuestion) {
            Row row = sheet.createRow(rowNum++);

            row.createCell(0)
                    .setCellValue(rowNum - 1);

            row.createCell(1)
                    .setCellValue(question.getDescription());

            row.createCell(2)
                    .setCellValue(question.getAnswerA());

            row.createCell(3)
                    .setCellValue(question.getAnswerB());

            row.createCell(4)
                    .setCellValue(question.getAnswerC());

            row.createCell(5)
                    .setCellValue(question.getAnswerD());

            row.createCell(6)
                    .setCellValue(question.getCorrectAnswer());
        }

        for (int i = 0; i < 6; i++) {
            sheet.autoSizeColumn(i);
        }

        FileOutputStream fileOut = new FileOutputStream(OUTPUT_EXCEL_PATH);
        workbook.write(fileOut);
        fileOut.close();
        workbook.close();
    }

    public static void writeExcel2(List<Question> listQuestion, String sheetName, String name) throws Exception {
        Workbook workbook;
        try {
            workbook = new XSSFWorkbook(new FileInputStream(OUTPUT_EXCEL_PATH));
        } catch (Exception ex) {
            workbook = new XSSFWorkbook();
        }

        //Generate path
        File path = new File(name.substring(0, name.lastIndexOf("/")));
        if (!path.exists()) {
            path.mkdir();
        }

        Sheet sheet = workbook.getSheetAt(0);

        int rowNum = 7;

        Font font = workbook.createFont();
        font.setBold(true);
        CellStyle styleBold = workbook.createCellStyle();
        styleBold.setFont(font);
        styleBold.setWrapText(true);

        CellStyle styleWrapText = workbook.createCellStyle();
        styleWrapText.setWrapText(true);

        for (Question question : listQuestion) {
            Row row = sheet.getRow(rowNum++);
            System.out.println(row.getCell(0));

            row.getCell(1).setCellStyle(styleBold);
            row.getCell(1)
                    .setCellValue(question.getDescription());

            row.createCell(9)
                    .setCellValue(question.getCorrectAnswer());

            System.out.println(row.getCell(1));

            row = sheet.getRow(rowNum++);

            row.getCell(1)
                    .setCellValue(question.getAnswerA());

            System.out.println(row.getCell(1));

            row.getCell(6)
                    .setCellValue(question.getAnswerB());

            System.out.println(row.getCell(4));

            row = sheet.getRow(rowNum++);

            row.getCell(1)
                    .setCellValue(question.getAnswerC());

            row.getCell(6)
                    .setCellValue(question.getAnswerD());
        }

        workbook.setSheetName(0, sheetName);
        FileOutputStream fileOut = new FileOutputStream(name);
        workbook.write(fileOut);
        fileOut.close();
        workbook.close();
    }

    public static void writeExce3(List<Question> listQuestion, String sheetName, String name) throws Exception {
        Workbook workbook;
        try {
            workbook = new XSSFWorkbook(new FileInputStream(OUTPUT_EXCEL_PATH));
        } catch (Exception ex) {
            workbook = new XSSFWorkbook();
        }

        //Generate path
        File path = new File(name.substring(0, name.lastIndexOf("/")));
        if (!path.exists()) {
            path.mkdir();
        }

        Sheet sheet = workbook.getSheetAt(0);

        int rowNum = 1;

        Font font = workbook.createFont();
        font.setBold(true);
        CellStyle styleBold = workbook.createCellStyle();
        styleBold.setFont(font);
        //styleBold.setWrapText(true);

        CellStyle styleWrapText = workbook.createCellStyle();
        styleWrapText.setWrapText(true);

        CellStyle styleBackgroundColorYellow = workbook.createCellStyle();
        styleBackgroundColorYellow.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
        styleBackgroundColorYellow.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        for (Question question : listQuestion) {
            Row row = sheet.createRow(rowNum++);
            setCellValueStyle(row.createCell(0), question.getDescription(), styleBold);
            System.out.println(row.getCell(0));

            for (int i = 0; i < 4; i++) {
                row = sheet.createRow(rowNum++);
                Cell cell = row.createCell(1);
                Cell cellPoint = row.createCell(2);

                if (i == 0 && question.getCorrectAnswer().equals("A")) {
                    setCellValueStyle(cell, question.answerA, styleBackgroundColorYellow);
                    cellPoint.setCellValue(1);
                    continue;
                } else if (i == 0) {
                    cell.setCellValue(question.getAnswerA());
                    cellPoint.setCellValue(0);
                    continue;
                }

                if (i == 1 && question.getCorrectAnswer().equals("B")) {
                    setCellValueStyle(cell, question.answerB, styleBackgroundColorYellow);
                    cellPoint.setCellValue(1);
                    continue;
                } else if (i == 1) {
                    cell.setCellValue(question.getAnswerB());
                    cellPoint.setCellValue(0);
                    continue;
                }

                if (i == 2 && question.getCorrectAnswer().equals("C")) {
                    setCellValueStyle(cell, question.answerC, styleBackgroundColorYellow);
                    cellPoint.setCellValue(1);
                    continue;
                } else if (i == 2) {
                    cell.setCellValue(question.getAnswerC());
                    cellPoint.setCellValue(0);
                    continue;
                }

                if (i == 3 && question.getCorrectAnswer().equals("D")) {
                    setCellValueStyle(cell, question.answerD, styleBackgroundColorYellow);
                    cellPoint.setCellValue(1);
                    continue;
                } else if (i == 3) {
                    cell.setCellValue(question.getAnswerD());
                    cellPoint.setCellValue(0);
                    continue;
                }
            }
        }

        workbook.setSheetName(0, sheetName);
        FileOutputStream fileOut = new FileOutputStream(name);
        workbook.write(fileOut);
        fileOut.close();
        workbook.close();
        System.out.println("Generated " + name);
    }

    private static void setCellValueStyle(Cell cell, String value, CellStyle cellStyle) {
        cell.setCellStyle(cellStyle);
        cell.setCellValue(value);
    }
}
