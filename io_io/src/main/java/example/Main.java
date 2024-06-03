package example;

import org.json.JSONObject;
import java.io.FileWriter;
import java.io.IOException;
import java.io.FileReader;


public class Main extends Data_table {
    public static void main(String[] args) {
        Data_table data_table = new Data_table();


        data_table.readFromExcel("_IO__.xlsx");


        JSONObject json = data_table.to_Json();
        System.out.println("JSON: " + json.toString());


        writeJsonToFile(json, "_IO_.json");


        //JSONObject jsonFromFile = readJsonFromFile("_IO_.json");
        //data_table.writeToExcelFromJson(jsonFromFile, "IO_IO.xlsx");
    }

}
