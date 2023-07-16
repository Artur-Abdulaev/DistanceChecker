import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
public class DistanceData implements Comparable<DistanceData> {
    public final OfficeAddress Actual;
    public final OfficeAddress Potential;
    public final double Distance;
    public final int Postcode;


    public static void main(String[] args) throws IOException {


        ArrayList<OfficeAddress> potential = new ArrayList<>();
        ArrayList<OfficeAddress> actual = new ArrayList<>();
        ArrayList<DistanceData> calculatedData = new ArrayList<>();

        FileInputStream potentialInputStream = new FileInputStream("C:/Users/A/Desktop/Programming/ExcelFiles/DistanceCalculator/ActualOffices.xlsx");
        XSSFWorkbook potentialWorkbookForRead = new XSSFWorkbook(potentialInputStream);
        XSSFSheet potentialSheetReadBook = potentialWorkbookForRead.getSheetAt(0);

        int potentialRowsCount = potentialSheetReadBook.getLastRowNum();

        for (int i = 1; i <= potentialRowsCount; i++) {
            String address = potentialSheetReadBook.getRow(i).getCell(0).getStringCellValue();
            double latitude = potentialSheetReadBook.getRow(i).getCell(1).getNumericCellValue();
            double longitude = potentialSheetReadBook.getRow(i).getCell(2).getNumericCellValue();
            potential.add(new OfficeAddress(address, latitude, longitude));
        }


        FileInputStream actualInputStreamActual = new FileInputStream("C:/Users/A/Desktop/Programming/ExcelFiles/DistanceCalculator/PPWZ.xlsx");
        XSSFWorkbook actualWorkbookForRead = new XSSFWorkbook(actualInputStreamActual);
        XSSFSheet actualSheetReadBook = actualWorkbookForRead.getSheetAt(0);

        int actualRowsCount = actualSheetReadBook.getLastRowNum();

        for (int i = 1; i <= actualRowsCount; i++) {
            String address = actualSheetReadBook.getRow(i).getCell(0).getStringCellValue();
            double latitude = actualSheetReadBook.getRow(i).getCell(1).getNumericCellValue();
            double longitude = actualSheetReadBook.getRow(i).getCell(2).getNumericCellValue();
            int postCode = (int) actualSheetReadBook.getRow(i).getCell(3).getNumericCellValue();
            actual.add(new OfficeAddress(address, latitude, longitude, postCode));
        }


        for (int i = 0; i < potential.size(); i++) {
            double min = Integer.MAX_VALUE;
            int minElementIndexPotential = -1;
            int minElementIndexActual = -1;
            for (int j = 0; j < actual.size(); j++) {
                if (getDistanceBetween(potential.get(i), actual.get(j)) < min) {
                    min = getDistanceBetween(potential.get(i), actual.get(j));
                    minElementIndexPotential = i;
                    minElementIndexActual = j;
                }
            }
            calculatedData.add(new DistanceData(potential.get(minElementIndexPotential), actual.get(minElementIndexActual), min, actual.get(minElementIndexActual).postCode));
        }

        Object[][] dataForWrite = Converter.convertTo2DArray(calculatedData);

        try {
            ExcelWriter.writer(dataForWrite, "C:/Users/�/Desktop/Programming/ExcelFiles/DistanceCalculator/Results.xlsx");
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }


    @Override
    public String toString() {
        return Distance + " ( " + Actual.toString() + " : " + Potential.toString() + " )\n";
    }

    public DistanceData(OfficeAddress actual, OfficeAddress potential, double distance, int postcode) {
        this.Actual = actual;
        this.Potential = potential;
        this.Distance = distance;
        this.Postcode = postcode;

    }

    public static double getDistanceBetween(OfficeAddress a1, OfficeAddress a2) {
        double theta = a1.Longitude - a2.Longitude;
        double lat1 = a1.Latitude;
        double lat2 = a2.Latitude;

        double dist = Math.sin(Math.toRadians(lat1)) * Math.sin(Math.toRadians(lat2)) +
                Math.cos(Math.toRadians(lat1)) * Math.cos(Math.toRadians(lat2)) * Math.cos(Math.toRadians(theta));

        dist = Math.acos(dist);
        dist = Math.toDegrees(dist);
        dist = dist * 60 * 1.1515;
        dist = dist * 1.609344 * 1000;

        return dist;
    }

    @Override
    public int compareTo(DistanceData o) {
        return this.Distance > o.Distance ? 1 : -1;
    }
}

