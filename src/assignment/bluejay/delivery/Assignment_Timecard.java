package assignment.bluejay.delivery;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.NumberToTextConverter;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Arrays;
import java.util.Date;
import java.util.Iterator;
import java.util.List;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

public class Assignment_Timecard {
    public static void main(String[] args) throws InvalidFormatException {
        try {
            FileInputStream file = new FileInputStream(new File("D:\\Assignment_Timecard.xlsx"));
            Workbook workbook = WorkbookFactory.create(file);

            // Assuming the data is in the first sheet
            Sheet sheet = workbook.getSheetAt(0);

            // Iterate through each row in the sheet
            Iterator<Row> rowIterator = sheet.iterator();
            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();

                // Assuming employee name is in the first column (index 0) and position is in the second column (index 1)
                String employeeName = getFormattedCellValue(row.getCell(7));
                String position = getFormattedCellValue(row.getCell(0));

                // Check the conditions
                if (hasWorkedFor7ConsecutiveDays(sheet, row) || hasLessThan20HoursBetweenShifts(row) || hasWorkedMoreThan14Hours(row)) {
                    System.out.println("Employee Name: " + employeeName + ", Position: " + position);
                }
            }

            workbook.close();
            file.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static boolean hasWorkedFor7ConsecutiveDays(Sheet sheet, Row currentRow) {
        Cell currentDateCell = currentRow.getCell(2);
        if(currentDateCell.toString().equals("")||currentDateCell.toString().equals("Time"))return false;
        if (currentDateCell != null) {
            String currentDateValue = getFormattedCellValue(currentDateCell);
            try {
                if (currentDateValue.equals("Time")||currentDateValue.equals("")) {
                // Handle the case where the value is "Time"
                return false;
            }
                Date currentDate = new SimpleDateFormat("MM/dd/yyyy hh:mm a").parse(currentDateValue);

                // Iterate through the previous 6 rows (6 days) and check if they are consecutive
                for (int i = 1; i <= 6; i++) {
                    Row previousRow = sheet.getRow(currentRow.getRowNum() - i);
                    if (previousRow != null) {
                        Cell previousDateCell = previousRow.getCell(2);
                         if(previousDateCell.toString().equals(""))return false;
                        if (previousDateCell != null) {
                            String previousDateValue = getFormattedCellValue(previousDateCell);
                            Date previousDate = new SimpleDateFormat("MM/dd/yyyy hh:mm a").parse(previousDateValue);
                            if (!isConsecutiveDates(previousDate, currentDate, i)) {
                                return false;
                            }
                        }
                    }
                }
                return true;
            } catch (ParseException e) {
                 System.out.println("problem in 7 consecutive days");
                e.printStackTrace();
            }
        }
        return false;
    }

    private static boolean hasLessThan20HoursBetweenShifts(Row row) {
        Cell timeInCell = row.getCell(5);
        Cell timeOutCell = row.getCell(6);
        
        if(timeInCell.toString().equals("")||timeOutCell.toString().equals("")||timeInCell.toString().equals("Pay Cycle Start Date")||timeOutCell.toString().equals("Pay Cycle End Date"))
        {
            return false;
        }
        if (timeInCell != null && timeOutCell != null) {
            String timeInValue = getFormattedCellValue(timeInCell);
            String timeOutValue = getFormattedCellValue(timeOutCell);
            try {
                Date timeIn = parseDate(timeInValue);
                Date timeOut = parseDate(timeOutValue);
                long timeDifference = Math.abs(timeOut.getTime() - timeIn.getTime());

                return timeDifference > 1 * 60 * 60 * 1000 && timeDifference < 20 * 60 * 60 * 1000;
            } catch (ParseException e) {
                 System.out.println("problem in 20 hours");
                e.printStackTrace();
            }
        }
        return false;
    }
    private static double convertToTotalHours(String timeString) {
        String[] timeParts = timeString.split(":");
        if (timeParts.length == 2) {
            int hours = Integer.parseInt(timeParts[0]);
            int minutes = Integer.parseInt(timeParts[1]);

            // Calculate total hours
            double totalHours = hours + (double) minutes / 60;
            return totalHours;
        } else {
            throw new IllegalArgumentException("Invalid time format: " + timeString);
        }
    }
    private static boolean hasWorkedMoreThan14Hours(Row row) {
        Cell timecardHoursCell = row.getCell(4);
        if (timecardHoursCell.toString().equals("Timecard Hours (as Time)")||timecardHoursCell.toString().equals("")) {
                // Handle the case where the value is "Timecard Hours (as Time)"
                return false;
            }
        if (timecardHoursCell != null) {    
            try {
                double timecardHours = convertToTotalHours(timecardHoursCell.toString());
                return timecardHours > 14;
            } catch (NumberFormatException e) {
                System.out.println("problem in 14 hours");
                e.printStackTrace();
            }
        }
        return false;
    }

    private static boolean isConsecutiveDates(Date previousDate, Date currentDate, int daysDifference) {
        long oneDayInMillis = 24 * 60 * 60 * 1000;
        return (currentDate.getTime() - previousDate.getTime()) == (daysDifference * oneDayInMillis);
    }
 private static Date parseDate(String dateValue) throws ParseException {
        // Updated list of date formats
        List<String> DATE_FORMATS = Arrays.asList(
                "MM/dd/yyyy",  // Updated format to cover the provided date
                "MM/dd/yyyy hh:mm a",
                "MM/dd/yyyy hh:mm:ss a",
                "yyyy-MM-dd'T'HH:mm:ss",
                "dd/MM/yyyy hh:mm a"
        );

        for (String format : DATE_FORMATS) {
            try {
                return new SimpleDateFormat(format).parse(dateValue);
            } catch (ParseException ignored) {
                // Continue to the next format if parsing fails
            }
        }
        throw new ParseException("Unable to parse date: " + dateValue, 0);
    }
    private static String getFormattedCellValue(Cell cell) {
        DataFormatter dataFormatter = new DataFormatter();
        return dataFormatter.formatCellValue(cell);
    }
}