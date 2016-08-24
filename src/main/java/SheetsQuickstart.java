import com.google.api.client.util.DateTime;
import com.google.api.services.sheets.v4.model.*;
import com.google.api.services.sheets.v4.Sheets;

import java.io.IOException;
import java.util.LinkedList;
import java.util.List;

public class SheetsQuickstart {
  public static final String SPREADSHEET_ID = "15gZhgEnynebN8_IMrCiXL2QFIMPq2OnC8Ni0eVuy0Po";
  public static final String CURRENT_RANGE_LOCATION = "C2:C2";

  public static void main(String[] args) throws IOException {
    Sheets service = SheetsHelper.getSheetsService();

    List<List<Object>> values = SheetsHelper.getValues(service, SPREADSHEET_ID, CURRENT_RANGE_LOCATION);
    String writeRange = values.get(0).get(0).toString();

    Request appendRequest = new Request();
    appendRequest.set("ValueInputOption", "USER_ENTERED");

    List<List<Object>> data = new LinkedList<>();
    LinkedList<Object> row1 = new LinkedList<>();
    row1.add("Reply Randy");
    row1.add(new DateTime("2016-08-22T11:59:00.00Z"));
    data.add(row1);
    ValueRange toWrite = new ValueRange();
    toWrite.setMajorDimension("ROWS");
    toWrite.setRange(writeRange);
    toWrite.setValues(data);
    AppendValuesResponse res = service.spreadsheets().values()
            .append(SPREADSHEET_ID, writeRange, toWrite)
            .set("valueInputOption", "USER_ENTERED")
            .execute();

    System.out.println(res.get("updates"));
    writeRange = updateWriteRange(writeRange);

    data = new LinkedList<>();
    row1 = new LinkedList<>();
    row1.add(writeRange);
    data.add(row1);
    toWrite = new ValueRange();
    toWrite.setMajorDimension("ROWS");
    toWrite.setRange("C2:C2");
    toWrite.setValues(data);
    service.spreadsheets().values()
         .update(SPREADSHEET_ID, "C2:C2", toWrite)
         .set("valueInputOption", "USER_ENTERED")
         .execute();
  }

  public static String updateWriteRange(String writeRange){
    int aVal = Integer.parseInt(writeRange.substring(1, writeRange.indexOf(':')));
    int bVal = Integer.parseInt(writeRange.substring(writeRange.indexOf(':')+2), writeRange.length());
    System.out.printf("%d, %d\n", aVal, bVal);
    return "" + writeRange.charAt(0) + (aVal+1) + ':' + writeRange.charAt(writeRange.indexOf(':')+1) + (bVal+1);
  }
}
