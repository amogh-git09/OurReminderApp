import com.google.api.client.auth.oauth2.Credential;
import com.google.api.client.extensions.java6.auth.oauth2.AuthorizationCodeInstalledApp;
import com.google.api.client.extensions.jetty.auth.oauth2.LocalServerReceiver;
import com.google.api.client.googleapis.auth.oauth2.GoogleAuthorizationCodeFlow;
import com.google.api.client.googleapis.auth.oauth2.GoogleClientSecrets;
import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
import com.google.api.client.http.HttpTransport;
import com.google.api.client.json.jackson2.JacksonFactory;
import com.google.api.client.json.JsonFactory;
import com.google.api.client.util.DateTime;
import com.google.api.client.util.store.FileDataStoreFactory;
import com.google.api.services.sheets.v4.SheetsScopes;
import com.google.api.services.sheets.v4.model.*;
import com.google.api.services.sheets.v4.Sheets;

import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.util.Arrays;
import java.util.LinkedList;
import java.util.List;

public class SheetsQuickstart {
  private static final String APPLICATION_NAME = "Google Sheets API Java Quickstart";
  private static final java.io.File DATA_STORE_DIR = new java.io.File(System.getProperty("user.home"), ".credentials/sheets.googleapis.com-java-quickstart");
  private static FileDataStoreFactory DATA_STORE_FACTORY;
  private static final JsonFactory JSON_FACTORY = JacksonFactory.getDefaultInstance();
  private static HttpTransport HTTP_TRANSPORT;
  private static final List<String> SCOPES = Arrays.asList(SheetsScopes.SPREADSHEETS);

  static {
    try {
      HTTP_TRANSPORT = GoogleNetHttpTransport.newTrustedTransport();
      DATA_STORE_FACTORY = new FileDataStoreFactory(DATA_STORE_DIR);
    } catch (Throwable t) {
      t.printStackTrace();
      System.exit(1);
    }
  }

  public static Credential authorize() throws IOException {
    //Load client secrets
    InputStream in = SheetsQuickstart.class.getResourceAsStream("/client_secret.json");
    GoogleClientSecrets clientSecrets = GoogleClientSecrets.load(JSON_FACTORY, new InputStreamReader(in));

    GoogleAuthorizationCodeFlow flow =
      new GoogleAuthorizationCodeFlow.Builder(HTTP_TRANSPORT, JSON_FACTORY, clientSecrets, SCOPES)
      .setDataStoreFactory(DATA_STORE_FACTORY)
      .setAccessType("offline")
      .build();

    Credential credential = new AuthorizationCodeInstalledApp(
        flow, new LocalServerReceiver()).authorize("user");

    System.out.println("Credentials saved to " + DATA_STORE_DIR.getAbsolutePath());
    return credential;
  }

  public static Sheets getSheetsService() throws IOException {
    Credential credential = authorize();
    return new Sheets.Builder(HTTP_TRANSPORT, JSON_FACTORY, credential)
        .setApplicationName(APPLICATION_NAME)
        .build();
  }

  public static void main(String[] args) throws IOException {
    Sheets service = getSheetsService();
    String spreadsheetId = "15gZhgEnynebN8_IMrCiXL2QFIMPq2OnC8Ni0eVuy0Po";
    String range = "C2:C2";
    ValueRange response = service.spreadsheets().values()
          .get(spreadsheetId, range)
          .execute();
    List<List<Object>> values = response.getValues();
    String writeRange = "";
    if(values == null || values.size() == 0){
      System.out.println("No data found.");
    } else {
      for(List row: values) {
        writeRange = row.get(0).toString();
        System.out.printf("%s\n", row.get(0));
      }
    }

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
            .append(spreadsheetId, writeRange, toWrite)
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
         .update(spreadsheetId, "C2:C2", toWrite)
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
