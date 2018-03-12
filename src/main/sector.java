package main;

import java.io.BufferedReader;
import java.io.DataInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileReader;
import java.io.IOException;
import java.io.Reader;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.Scanner;

import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVParser;
import org.apache.commons.csv.CSVRecord;
import org.apache.log4j.Logger;

public class sector {
	
	static Logger logger = Logger.getLogger(sector.class);
    public ArrayList<String> sectorValues() {
    	ArrayList<String> result=new ArrayList<String>();
        // -define .csv file in app
        String fileNameDefined = "ShareNumber.csv";
        // -File class needed to turn stringName to actual file
        //File file = new File(fileNameDefined);
        String thisLine=null;
        try{
            // -read from filePooped with Scanner class
        	//Scanner sc= new Scanner(System.in);
        	//String line = "";
        	//int lineNumber=sc.nextInt();
        	Reader br = Files.newBufferedReader(Paths.get(fileNameDefined));
        	CSVParser csvParser = new CSVParser(br,CSVFormat.DEFAULT);
        	Iterable<CSVRecord> csvRecords = csvParser.getRecords();
            for(CSVRecord csvRecord: csvRecords){
                //read single line, put in string
                String data = csvRecord.get(1);
                result.add(data);

            }
            // after loop, close scanner
            br.close();


        }catch (IOException e){

        	logger.fatal("exception at reading", e);
        }
        return result;

    }

}
