package DatabaseQueryExecutor.DQE;

import javafx.application.Application;
import javafx.geometry.Insets;
import javafx.geometry.Pos;
import javafx.scene.Scene;
import javafx.scene.control.Button;
import javafx.scene.control.DatePicker;
import javafx.scene.control.Label;
import javafx.scene.layout.GridPane;
import javafx.stage.Stage;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.Statement;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.JOptionPane;


public class App extends Application {

    @Override
    public void start(Stage primaryStage) {
        primaryStage.setTitle("SQL Query Executor");

        // Labels
        Label fromDateLabel = new Label("From Date:");
        Label toDateLabel = new Label("To Date:");

        // DatePickers
        DatePicker fromDatePicker = new DatePicker();
        DatePicker toDatePicker = new DatePicker();

        // Button
        Button validateButton = new Button("Validate");

        // GridPane Layout
        GridPane gridPane = new GridPane();
        gridPane.setAlignment(Pos.CENTER);
        gridPane.setPadding(new Insets(10));
        gridPane.setHgap(10);
        gridPane.setVgap(10);

        // Add nodes to GridPane
        gridPane.add(fromDateLabel, 0, 0);
        gridPane.add(fromDatePicker, 1, 0);
        gridPane.add(toDateLabel, 0, 1);
        gridPane.add(toDatePicker, 1, 1);
        gridPane.add(validateButton, 1, 2);

        // Button Action
        validateButton.setOnAction(e -> {
            // Add your database connection and query execution code here
            String fromDate = fromDatePicker.getValue().toString();
            String toDate = toDatePicker.getValue().toString();
            DatabaseQueryExecutor.runQuery(fromDate, toDate);
        });

        // Set Scene
        Scene scene = new Scene(gridPane, 400, 200);
        primaryStage.setScene(scene);
        primaryStage.show();
    }

    public static void main(String[] args) {
        launch(args);
    }
    
    public class DatabaseQueryExecutor {

        public static void runQuery(String fromDate, String toDate) {
            String excelFilePath = "src/resources/query.xlsx";
            String query = readQueryFromExcel(excelFilePath);

            if (query == null) {
                JOptionPane.showMessageDialog(null, "Failed to read the query from Excel.");
                return;
            }

            query = query.replace(":fromDate", fromDate).replace(":toDate", toDate);

            try (Connection connection = DriverManager.getConnection("jdbc:sqlserver://localhost;integratedSecurity=true")) {
                Statement statement = connection.createStatement();
                ResultSet resultSet = statement.executeQuery(query);

                String outputFilePath = "C:/Users/Reports/QueryResults.xlsx";
                writeResultsToExcel(resultSet, outputFilePath);

                JOptionPane.showMessageDialog(null, "Query Exported and saved into folder: " + outputFilePath);
            } catch (Exception e) {
                e.printStackTrace();
                JOptionPane.showMessageDialog(null, "An error occurred: " + e.getMessage());
            }
        }

        private static String readQueryFromExcel(String filePath) {
            try (FileInputStream fis = new FileInputStream(new File(filePath));
                 Workbook workbook = new XSSFWorkbook(fis)) {

                Sheet sheet = workbook.getSheetAt(0);
                Row row = sheet.getRow(0);
                Cell cell = row.getCell(0);
                return cell.getStringCellValue();

            } catch (Exception e) {
                e.printStackTrace();
                return null;
            }
        }

        private static void writeResultsToExcel(ResultSet resultSet, String outputPath) {
            try (Workbook workbook = new XSSFWorkbook(); 
                 FileOutputStream fos = new FileOutputStream(outputPath)) {

                Sheet sheet = workbook.createSheet("Query Results");

                Row headerRow = sheet.createRow(0);
                int columnCount = resultSet.getMetaData().getColumnCount();

                for (int i = 1; i <= columnCount; i++) {
                    headerRow.createCell(i - 1).setCellValue(resultSet.getMetaData().getColumnName(i));
                }

                int rowCount = 1;

                while (resultSet.next()) {
                    Row row = sheet.createRow(rowCount++);
                    for (int i = 1; i <= columnCount; i++) {
                        row.createCell(i - 1).setCellValue(resultSet.getString(i));
                    }
                }

                workbook.write(fos);
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
    }
}
