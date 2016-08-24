import com.google.api.client.auth.oauth2.Credential;
import com.google.api.client.extensions.java6.auth.oauth2.AuthorizationCodeInstalledApp;
import com.google.api.client.extensions.jetty.auth.oauth2.LocalServerReceiver;
import com.google.api.client.googleapis.auth.oauth2.GoogleAuthorizationCodeFlow;
import com.google.api.client.googleapis.auth.oauth2.GoogleClientSecrets;
import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
import com.google.api.client.http.HttpTransport;
import com.google.api.client.json.JsonFactory;
import com.google.api.client.json.jackson2.JacksonFactory;
import com.google.api.client.util.store.FileDataStoreFactory;
import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.SheetsScopes;
import com.google.api.services.sheets.v4.model.UpdateValuesResponse;
import com.google.api.services.sheets.v4.model.ValueRange;

import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.util.Arrays;
import java.util.LinkedList;
import java.util.List;

/**
 * Created by amogh-lab on 16/08/24.
 */
public class SheetsHelper {
    private static final String APPLICATION_NAME = "Google Sheets API Java Quickstart";
    private static final java.io.File DATA_STORE_DIR = new java.io.File(System.getProperty("user.home"), ".credentials/sheets.googleapis.com-java-quickstart");
    private static FileDataStoreFactory DATA_STORE_FACTORY;
    private static final JsonFactory JSON_FACTORY = JacksonFactory.getDefaultInstance();
    private static HttpTransport HTTP_TRANSPORT;
    private static final List<String> SCOPES = Arrays.asList(SheetsScopes.SPREADSHEETS);

    private Sheets service;
    private String spreadSheetId;

    public SheetsHelper(Sheets service, String id){
        this.service = service;
        this.spreadSheetId = id;
    }

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

    public List<List<Object>> getValues(String loc) throws java.io.IOException{
        ValueRange response = service.spreadsheets().values()
                .get(spreadSheetId, loc)
                .execute();
        return response.getValues();
    }

    public UpdateValuesResponse writeToSheet(String range, List<List<Object>> data) throws java.io.IOException{
        ValueRange toWrite = new ValueRange();
        toWrite.setMajorDimension("ROWS");
        toWrite.setRange(range);
        toWrite.setValues(data);

        return service.spreadsheets().values()
                .update(spreadSheetId, range, toWrite)
                .set("valueInputOption", "USER_ENTERED")
                .execute();
    }

    public UpdateValuesResponse updateWriteRange(String writeRange) throws java.io.IOException{
        int aVal = Integer.parseInt(writeRange.substring(1, writeRange.indexOf(':')));
        int bVal = Integer.parseInt(writeRange.substring(writeRange.indexOf(':')+2, writeRange.length()));
        writeRange = "" + writeRange.charAt(0) + (aVal+1) + ':' + writeRange.charAt(writeRange.indexOf(':')+1) + (bVal+1);
        List<List<Object>> data = new LinkedList<>();
        List<Object> row1 = new LinkedList<>();
        row1.add(writeRange);
        data.add(row1);
        return writeToSheet(SheetsQuickstart.CURRENT_RANGE_LOCATION, data);
    }
}
