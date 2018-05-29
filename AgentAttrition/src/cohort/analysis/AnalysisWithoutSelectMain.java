package cohort.analysis;

import java.io.IOException;
import java.text.ParseException;

public class AnalysisWithoutSelectMain {
	public static void main(String args[]) throws ParseException, IOException
	{
		AnalysisWithoutSelect ana = new AnalysisWithoutSelect();
		String input[] = new String[2]; 
		input[0]="AGENT_ATTRITION_NO_SELECT_JAN17-DEC17.csv";
		input[1]="Agent_Attrition_No_Select_Results.xlsx";
		ana.main(input);
		input[0]="AGENT_ATTRITION_NO_SELECT_JAN18-APR18.csv";
		input[1]="Agent_Attrition_No_Select_Results1.xlsx";
		ana.main(input);
	}
}
