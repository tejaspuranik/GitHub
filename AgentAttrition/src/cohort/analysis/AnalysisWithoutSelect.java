package cohort.analysis;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.io.InputStreamReader;
import java.io.PrintWriter;
import java.text.DecimalFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.TimeZone;

import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class AnalysisWithoutSelect 
{      
	 //Just checking github setup
      static int[] totalVolume;
      
      static int[][] segQuartileCountMatrix;
      static int[][] segQuartileFailCountMatrix;
      
      static double[][] abtSumMatrix;
      static int[][] abtCountMatrix;
      static double[][] abtSumFailMatrix;
      static int[][] abtCountFailMatrix;
      
      static double[][] fcrSumMatrix;
      static int[][] fcrCountMatrix;
      static double[][] fcrSumFailMatrix;
      static int[][] fcrCountFailMatrix;
      
      static double[][] sv60SumMatrix;
      static int[][] sv60CountMatrix;
      static double[][] sv60SumFailMatrix;
      static int[][] sv60CountFailMatrix;
      
      static int[][] totalAgentsByCategory = new int[12][5];
      static int[][] totalAgentsByCategoryAttrition = new int[12][5];
      static int[][] AgentsAttritionByCategory = new int[12][5];
      static String site = null,partner =null,location =null,answer=null,lob=null,result=null,segmentName=null,queueName=null;
      
      static Sheet sheet,sheet2,sheet3,sheet4,sheet5,sheet6;
      static int rowCount = 0,rowCountSheet2=0,rowCountSheet3=0,rowCountSheet4=0,rowCountSheet5=0,rowCountSheet6=0;
      
      static boolean debugFlag=true,NormalizeFlag=false;
      static StringBuilder sb;
      
     enum Operation {
    	  QuartilePerformance,
    	  QPNormalized,
    	  Attrition;
      }
     private static DecimalFormat df2 = new DecimalFormat(".##");
      
	  public static void main(String[] args) throws ParseException, IOException {		  
	        String csvFile = "./"+args[0];//AGENT_FILE_JAVA_ATT_COHORT_V032.csv
	        BufferedReader br = null;
	        BufferedReader br1 = null;
	       	        
	        String line = "";	      
	        ArrayList<MyAgent> agents = new ArrayList<MyAgent>();
	        SimpleDateFormat formatter = new SimpleDateFormat("ddMMMyyyy");
	        formatter.setTimeZone(TimeZone.getTimeZone("GMT-04:00"));
	        
	        // Create a Workbook
	        Workbook workbook = new XSSFWorkbook();     // new HSSFWorkbook() for generating `.xls` file

	        /* CreationHelper helps us create instances for various things like DataFormat,
	           Hyperlink, RichTextString etc in a format (HSSF, XSSF) independent way */
	        CreationHelper createHelper = workbook.getCreationHelper();

	        // Create a Sheet
	        sheet = workbook.createSheet("KPI Analysis");
	        sheet2 = workbook.createSheet("Quartile Performance");
	        sheet3 = workbook.createSheet("Quartile Performance Normalized");
	        sheet4 = workbook.createSheet("Attrition");
	        sheet5 = workbook.createSheet("Total Volume");
	        sheet6 = workbook.createSheet("Attrition Without Quartiles");
	        
	        PrintWriter pw = new PrintWriter(new File("./"+"Debug.csv"));
	        sb = new StringBuilder();
	      
	        int segIndex=21;//Hardcoded for reading purposes change later
	        int numberOfMonths =12;
	        	        	        
	        br1 = new BufferedReader(new InputStreamReader(System.in));
	        System.out.println("Enter LOB");
        	lob = br1.readLine();
        		       
	        System.out.println("Enter y to specify partner or n to continue analysis");
	        answer = br1.readLine();
	        if(answer.equalsIgnoreCase("y"))
	        {
	        	System.out.println("Enter Partner");
	        	partner = br1.readLine();
	        }
        	 
	        System.out.println("Enter y to specify site or n to continue analysis");
	        answer = br1.readLine();
	        if(answer.equalsIgnoreCase("y"))
	        {
	        	System.out.println("Enter Site");
	        	site = br1.readLine();
	        }
	        System.out.println("Enter y to specify location or n to continue analysis");
	        answer = br1.readLine();
	        if(answer.equalsIgnoreCase("y"))
	        {
	        	 System.out.println("Enter Location");
	 	        location = br1.readLine();
	        }
	        System.out.println("Enter y to specify queue name or n to continue analysis");
	        answer = br1.readLine();
	        if(answer.equalsIgnoreCase("y"))
	        {
	        	 System.out.println("Enter Queue Name");
	        	 queueName = br1.readLine();
	        }
	        System.out.println("Enter y to specify segment name or n to continue analysis");
	        answer = br1.readLine();
	        if(answer.equalsIgnoreCase("y"))
	        {
	        	 System.out.println("Enter Segment Name");
	        	 segmentName = br1.readLine();
	        }
	        
	        /*System.out.println("Enter y to specify result or n to continue analysis");
	        answer = br1.readLine();
	        if(answer.equalsIgnoreCase("y"))
	        {
	        	 System.out.println("Enter Result");
	        	 result = br1.readLine();
	        }*/
	       	       	        	       
            boolean siteFlag= (site ==null);
            boolean locationFlag = (location ==null);
            boolean partnerFlag = (partner ==null);
            
            String[] removedNull,removedNullAbts,removedNullFcr,removedNullSv60, removedNullSegNm, removedNullQueNm;
            ArrayList<String[]> attrition = new ArrayList<String[]>();
            
            int skipLine=0;
                    	        
	        try 
	        {	        
	            br = new BufferedReader(new FileReader(csvFile));
	            
	            while ((line = br.readLine()) != null) 
	            {
	            	if(skipLine==0)
	            	{
	            		skipLine++;
	            		continue;
	            	}
	            		
	                // use comma as separator
	            	//line = line.replace("\"", "");
	                String[] fields = line.split(",(?=(?:[^\"]*\"[^\"]*\")*[^\"]*$)", -1);
	                for(int t1=0;t1<fields.length;t1++)
	                {
	                	fields[t1]=fields[t1].replaceAll("\"", "");
	                }
	                String[] segments = Arrays.copyOfRange(fields, segIndex, segIndex+numberOfMonths);
	                
	                removedNull = Arrays.stream(segments)
	                        .filter(value ->
	                                value != null && value.length() > 0
	                        )
	                        .toArray(size -> new String[size]);
	                attrition.add(removedNull);
	                
	                String[] abts = Arrays.copyOfRange(fields, segIndex+numberOfMonths, segIndex+(numberOfMonths*2));
	                removedNullAbts= Arrays.stream(abts)
	                        .filter(value ->
                            value != null && !value.equals(".") && value.length() > 0
                    ).toArray(size -> new String[size]);
	                
	                String[] fcr = Arrays.copyOfRange(fields, segIndex+(numberOfMonths*2), segIndex+(numberOfMonths*3));
	                removedNullFcr= Arrays.stream(fcr)
	                        .filter(value ->
                            value != null && !value.equals(".") && value.length() > 0
                    ).toArray(size -> new String[size]);
	                
	                String[] sv60 = Arrays.copyOfRange(fields, segIndex+(numberOfMonths*3), segIndex+(numberOfMonths*4));
	                removedNullSv60= Arrays.stream(sv60)
	                        .filter(value ->
                            value != null && !value.equals(".") && value.length() > 0
                    ).toArray(size -> new String[size]);
	                
	                String[] segNameMonths = Arrays.copyOfRange(fields, segIndex+(numberOfMonths*4), segIndex+(numberOfMonths*5)+1);//As there is one month 	                															
	                removedNullSegNm= Arrays.stream(segNameMonths)																	//extra for segment/Queue
	                        .filter(value ->
                            value != null && !value.equals(".") && value.length() > 0
                    ).toArray(size -> new String[size]);
	                
	                String[] queueNameMonths = Arrays.copyOfRange(fields, segIndex+(numberOfMonths*5)+1, segIndex+(numberOfMonths*6)+1);
	                removedNullQueNm= Arrays.stream(queueNameMonths)
	                        .filter(value ->
                            value != null && !value.equals(".") && value.length() > 0
                    ).toArray(size -> new String[size]);
	                
	                
	                MyAgent agent = new MyAgent();
	            	int i=0;            	
	            	
	            	ArrayList<Double> abtMonths = new ArrayList<Double>(12);
	            	for(int o=0;o<12;o++)
	            	{
	            		if(o<removedNullAbts.length)
	            			abtMonths.add(Double.parseDouble(removedNullAbts[o]));	   
	            		else
	            			abtMonths.add((double) 0);
	            	}
	            	agent.setAbtMonths(abtMonths);
	            	                     	            	
	            	ArrayList<Float> segmonths = new ArrayList<Float>(12);
	            	for(int o=0;o<removedNull.length;o++)
	            	{
	            		if(o<12)//we need first 12 months of consecutive data only
	            		segmonths.add(Float.parseFloat(removedNull[o]));	            		            		
	            	}
	            	agent.setSegMonths(segmonths);
	            	
	            	ArrayList<Double> fcrMonths = new ArrayList<Double>(12);
	            	for(int o=0;o<12;o++)
	            	{
	            		if(o<removedNullFcr.length)
	            			fcrMonths.add(Double.parseDouble(removedNullFcr[o]));	   
	            		else
	            			fcrMonths.add((double) 0);
	            	}
	            	agent.setFcrMonths(fcrMonths);
	            	
	            	ArrayList<Double> sv60Months = new ArrayList<Double>(12);
	            	for(int o=0;o<12;o++)
	            	{
	            		if(o<removedNullSv60.length)
	            			sv60Months.add(Double.parseDouble(removedNullSv60[o]));	   
	            		else
	            			sv60Months.add((double) 0);
	            	}
	            	agent.setSv60Months(sv60Months);
	            	
	            	ArrayList<String> segNmMonths = new ArrayList<String>(12);
	            	for(int o=0;o<12;o++)
	            	{
	            		if(o<removedNullSegNm.length)
	            			segNmMonths.add(removedNullSegNm[o]);	   
	            	}
	            	agent.setSegNameMonths(segNmMonths);
	            	
	            	ArrayList<String> queueNmMonths = new ArrayList<String>(12);
	            	for(int o=0;o<12;o++)
	            	{
	            		if(o<removedNullQueNm.length)
	            			queueNmMonths.add(removedNullQueNm[o]);	   
	            	}
	            	agent.setQueueNameMonths(queueNmMonths);
	            	
	                while(i<segIndex)
	                {
	                	switch(i)
	                	{
		                	case 0:
	                			agent.setEmpNo(fields[i]);
	                			break;
	                		case 1:
	                			agent.setHireDate(formatter.parse(fields[i]));
	                			break;
	                		case 2:
	                			agent.setPartner(fields[i]);       			
	                			break;
	                		case 3:
	                			agent.setSiteDesc(fields[i]);
	                			break;
	                		case 4:
	                			agent.setLocation(fields[i]);
	                			break;
	                		case 5:
	                			agent.setAttritionCount(Integer.parseInt(fields[i]));
	                			break;
	                		case 6:
	                			agent.setAttritionFlag(Integer.parseInt(fields[i]));
	                			break;
	                		case 7:
	                			agent.setJan31Diff(Integer.parseInt(fields[i]));
	                			break;
	                		case 8:
	                			agent.setFeb28Diff(Integer.parseInt(fields[i]));
	                			break;
	                		case 9:
	                			agent.setMar31Diff(Integer.parseInt(fields[i]));
	                			break;
	                		case 10:
	                			agent.setApr30Diff(Integer.parseInt(fields[i]));
	                			break;
	                		case 11:
	                			agent.setMay31Diff(Integer.parseInt(fields[i]));
	                			break;  
	                		case 12:
	                			agent.setJun30Diff(Integer.parseInt(fields[i]));
	                			break;   
	                		case 13:
	                			agent.setJul31Diff(Integer.parseInt(fields[i]));
	                			break;   
	                		case 14:
	                			agent.setAug31Diff(Integer.parseInt(fields[i]));
	                			break;   
	                		case 15:
	                			agent.setSep30Diff(Integer.parseInt(fields[i]));
	                			break;   
	                		case 16:
	                			agent.setOct31Diff(Integer.parseInt(fields[i]));
	                			break;  
	                		case 17:
	                			agent.setNov30Diff(Integer.parseInt(fields[i]));
	                			break; 
	                		case 18:
	                			agent.setDec31Diff(Integer.parseInt(fields[i]));
	                			break; 	                		
	                		case 19:
	                			agent.setMay01Diff(Integer.parseInt(fields[i]));
	                			break; 
	                		case 20:
	                			agent.setLob(fields[i]);	                			
	                			break; 
	                	}
	                	i++;
	                }	                	               
	                agents.add(agent);
	            }
	          
	            for(int w=0;w<1;w++)
	            {
	                abtSumMatrix = new double[12][4];
	    	        abtCountMatrix = new int[12][4];
	    	        abtSumFailMatrix = new double[12][4];
	    	        abtCountFailMatrix = new int[12][4];
	    	        
	    	        fcrSumMatrix = new double[12][4];
	    	        fcrCountMatrix = new int[12][4];
	    	        fcrSumFailMatrix = new double[12][4];
	    	        fcrCountFailMatrix = new int[12][4];
	    	        
	    	        sv60SumMatrix = new double[12][4];
	    	        sv60CountMatrix = new int[12][4];
	    	        sv60SumFailMatrix = new double[12][4];
	    	        sv60CountFailMatrix = new int[12][4];
	    	        
	    	        segQuartileCountMatrix = new int[12][4];
	    	        segQuartileFailCountMatrix = new int[12][4];

	    	        for(MyAgent agent:agents)
	 	            {   	 	            	
	 	            	if(agent.getLob().equalsIgnoreCase(lob))
	 	            	{	   
	 	            		if(!siteFlag)
	 						{						
	 	            			siteFlag = agent.getSiteDesc().equalsIgnoreCase(site);
	 						}
	 	            		if(!partnerFlag)
	 						{						
	 	            			partnerFlag = agent.getPartner().equalsIgnoreCase(partner);
	 						}
	 	            		if(!locationFlag)
	 						{							 	            			
	 	            			locationFlag = agent.getLocation().equalsIgnoreCase(location) && !agent.getPartner().equalsIgnoreCase(partner);
	 	            			partnerFlag=true;
	 						}
	 	            			 						
	 	            		if(siteFlag && locationFlag && partnerFlag)
	 	            		{	 	   
	 	            			quartilePerformanceKpiAnalysis(agent);
 	            			    quartilePerformanceListAnalysis(agent.getSegMonths(),agent,Operation.QuartilePerformance);
	 	            		}	
	 	                    siteFlag= (site ==null);
	 	                    locationFlag = (location ==null);
	 	                    partnerFlag = (partner ==null);
	 	            	}	            			            	
	 	            }
	            	 	    	      
	            	System.out.println("Quartile Peformance Per Kpi Analysis");
	 	            displayKpiAnalysis();
	 	            
	 	            System.out.println();
		            System.out.println("Quartile Peformance Analysis");
		            display(result,Operation.QuartilePerformance);

		            for(int r1=0;r1<totalAgentsByCategory.length;r1++)
		            {		       		            	
		            	totalAgentsByCategory[r1][totalAgentsByCategory[0].length-1] = Arrays.stream(segQuartileFailCountMatrix[r1]).sum();
		            	totalAgentsByCategory[r1][w] = Arrays.stream(segQuartileCountMatrix[r1]).sum();		            	
		            }
		          
		            segQuartileCountMatrix = new int[12][4];
	    	        segQuartileFailCountMatrix = new int[12][4];
		            
		            for(MyAgent agent:agents)
		            {
		            	if(agent.getMay01Diff()>0)
		            	{
		            		float sum = 0;
		            		float average =0;
			            	for(Float a:agent.getSegMonths()) 
			            	{
			            		sum+=a;
			            	}
			            	if(agent.getSegMonths().size()>0)
			            		average = (float) Math.floor(sum/agent.getSegMonths().size()); 
			            	int left = 12-agent.getSegMonths().size();
			            	ArrayList<Float> tempSeg = new ArrayList<Float>(agent.getSegMonths());
			            	
			            	for(int i=0;i<left;i++)
			            	{			   
			            		tempSeg.add(average);
			            	}
			            	agent.setSegProjectedMonths(tempSeg);
		            	}
		            }
		            
		            for(MyAgent agent:agents)
		            {   
		            	if(agent.getMay01Diff()>0 && agent.getLob().equalsIgnoreCase(lob))
		            	{
		            		if(!siteFlag)
	 						{						
	 	            			siteFlag = agent.getSiteDesc().equalsIgnoreCase(site);
	 						}
		            		if(!partnerFlag)
	 						{						
	 	            			partnerFlag = agent.getPartner().equalsIgnoreCase(partner);
	 						}
	 	            		if(!locationFlag)
	 						{							 	            			
	 	            			locationFlag = agent.getLocation().equalsIgnoreCase(location) && !agent.getPartner().equalsIgnoreCase(partner);
	 	            			partnerFlag=true;
	 						}
	 	            			 	            		
	 	            		if(siteFlag && locationFlag && partnerFlag)//	&& agent.getQueueNameMonths().get(0).equalsIgnoreCase(queueName) && agent.getSegNameMonths().get(0).equalsIgnoreCase(segmentName) 	 	         			
	 	            		{
	 	            			//NormalizeFlag=true;
	 	            			quartilePerformanceListAnalysis(agent.getSegProjectedMonths(),agent, Operation.QPNormalized);
	 	            			//NormalizeFlag=false;
	 	            		}	
	 	                    siteFlag= (site ==null);
	 	                    locationFlag = (location ==null);
	 	                    partnerFlag = (partner ==null);	 	                    		            	
		            	}	            		            	
		            }
		            	           
		            System.out.println("Quartile Peformance Projected Analysis");
		            display(result,Operation.QPNormalized);
		            
		            System.out.println();
		            
		            segQuartileCountMatrix = new int[12][4];
	    	        segQuartileFailCountMatrix = new int[12][4];
	    	        
	    	        int k1=0;
		            for(MyAgent agent:agents)
		            { 
		            	if(agent.getAttritionFlag()==1)
		            	{
		            		int monthNo = (int) Math.ceil(agent.getAttritionCount() /30);	            	

			                if(agent.getAttritionCount()%30>0)
			                	monthNo+=1;
			                ArrayList<Float> tempAttr = new ArrayList<Float>(12);
			                
			                for(int z=0;z<12;z++)
			                {
			                	tempAttr.add((float) 0);
			                }
			                
			                int last = attrition.get(k1).length-1;			                
			                
			                if(last>-1 && monthNo>0 && monthNo<13)
			                {
			                	//System.out.println(last+"\t"+monthNo+"\t"+attrition.get(k1).length+"\t"+agent.getAttritionCount()+"\t"+agent.getEmpNo());
			                	tempAttr.set(monthNo-1, Float.parseFloat(attrition.get(k1)[last]));
			                }
			              		              
			                agent.setSegAttritionMonths(tempAttr);
		            	}
		            	k1++;
		            }
		            		         		           		            
		            for(MyAgent agent:agents)
		            {   
		            	if(agent.getAttritionFlag()==1 && agent.getLob().equalsIgnoreCase(lob))
		            	{
		            		if(!siteFlag)
	 						{						
	 	            			siteFlag = agent.getSiteDesc().equalsIgnoreCase(site);
	 						}
		            		if(!partnerFlag)
	 						{						
	 	            			partnerFlag = agent.getPartner().equalsIgnoreCase(partner);
	 						}
	 	            		if(!locationFlag)
	 						{							 	            			
	 	            			locationFlag = agent.getLocation().equalsIgnoreCase(location) && !agent.getPartner().equalsIgnoreCase(partner);
	 	            			partnerFlag=true;
	 						}
	 	            		if(siteFlag && locationFlag && partnerFlag)
	 	            		{
	 	            			quartilePerformanceListAnalysis(agent.getSegAttritionMonths(),agent,Operation.Attrition);
	 	            			int monthNo = (int) Math.ceil(agent.getAttritionCount() /30);	            	
				                if(agent.getAttritionCount()%30>0)
				                	monthNo+=1;
				                
				                if(monthNo<13)
	 	            			attritionWithoutQuartileAnalysis(monthNo,agent);
	 	            		}	
	 	                    siteFlag= (site ==null);
	 	                    locationFlag = (location ==null);
	 	                    partnerFlag = (partner ==null);	 		 	                    	            		
		            	}	            	          	
		            }
		            System.out.println("Quartile Attrition Analysis");
		            //display(result,Operation.Attrition);
		            displayAttrition();
		         		           
		            totalVolume = new int[12];
		            for(MyAgent agent:agents)
		            {   
		            	if(agent.getLob().equalsIgnoreCase(lob))
		            	{
		            		if(!siteFlag)
	 						{						
	 	            			siteFlag = agent.getSiteDesc().equalsIgnoreCase(site);
	 						}
		            		if(!partnerFlag)
	 						{						
	 	            			partnerFlag = agent.getPartner().equalsIgnoreCase(partner);
	 						}
	 	            		if(!locationFlag)
	 						{							 	            			
	 	            			locationFlag = agent.getLocation().equalsIgnoreCase(location) && !agent.getPartner().equalsIgnoreCase(partner);
	 	            			partnerFlag=true;
	 						}
	 	            		if(siteFlag && locationFlag && partnerFlag)
	 	            		{
	 	            			attritionAnalysis(agent.getSegAttritionMonths(),agent);
	 	            			calculateTotalVolume(agent);
	 	            		}	
	 	                    siteFlag= (site ==null);
	 	                    locationFlag = (location ==null);
	 	                    partnerFlag = (partner ==null);	 		 	                    		            	
		            	}                       	
		            }
		            
		            System.out.println("Total Volume");	 	 		            
		            for(int e=0;e<totalVolume.length;e++)
		            {
		            	System.out.println(totalVolume[e]);
		            }
	            }
	            
	           displayTotalVolume(); 
 	           displayAttritionAnalysisWithoutQuartiles();
 	           initializeStaticVariables();
 	           // Write the output to a file
 	           FileOutputStream fileOut = new FileOutputStream("./"+args[1]);
 	           workbook.write(fileOut);
 	           fileOut.close(); 	           
	           	          
 	          workbook.close();
 	          pw.write(sb.toString()); 
 	          pw.close();
	        }
	        catch (FileNotFoundException e) {
	            e.printStackTrace();
	        } catch (IOException e) {
	            e.printStackTrace();
	        } finally {
	            if (br != null) {
	                try {
	                    br.close();
	                } catch (IOException e) {
	                    e.printStackTrace();
	                }
	            }
	        }
		}
		
		public static void quartilePerformanceListAnalysis(ArrayList<Float> list, MyAgent agent, Operation op)
		{
			float epsilon=(float) 0.00000001;
			int monthNo=-1,monthForQS=-1;
			switch(op)
			{
				case QuartilePerformance:
					break;
				case QPNormalized:
					break;												 
				case Attrition:
					for(int i1=0;i1<list.size();i1++)
					{
						if(Math.abs(list.get(i1) - 0) > epsilon)
						{
							monthNo=i1;
						}
					}
					break;
			}
						
			boolean segementFlag= segmentName==null;
			
			boolean queueFlag= queueName==null;
						
				for(int y=0;y<list.size();y++)
				{
					if(op==Operation.Attrition)
					{
						if(monthNo == y)
						{
							if(!segementFlag)
							{			
								monthForQS=monthNo;
								if(monthNo>=agent.getSegNameMonths().size())
								{
									monthForQS=agent.getSegNameMonths().size()-1;
								}
								segementFlag = agent.getSegNameMonths().get(monthForQS).equalsIgnoreCase(segmentName);
							}
							if(!queueFlag)
							{
								monthForQS=monthNo;
								if(monthNo>=agent.getQueueNameMonths().size())
								{
									monthForQS=agent.getQueueNameMonths().size()-1;
								}					
								queueFlag = agent.getQueueNameMonths().get(monthForQS).equalsIgnoreCase(queueName);
							}
						}						
					}
					else
					{
						if(!segementFlag && y<agent.getSegNameMonths().size())
						{						
							segementFlag = agent.getSegNameMonths().get(y).equalsIgnoreCase(segmentName);
						}
						if(!queueFlag && y<agent.getQueueNameMonths().size())
						{
							queueFlag = agent.getQueueNameMonths().get(y).equalsIgnoreCase(queueName);
						}
					}
															
					if((segementFlag && queueFlag) || NormalizeFlag)
					{
						if(Math.abs(list.get(y) - 1) < epsilon)
		            	{
							segQuartileCountMatrix[y][0]+= 1;
		            	}
						if(Math.abs(list.get(y) - 2) < epsilon)
		            	{
							segQuartileCountMatrix[y][1]+= 1;
		            	}
						if(Math.abs(list.get(y) - 3) < epsilon)
		            	{
							segQuartileCountMatrix[y][2]+= 1;
		            	}
						if(Math.abs(list.get(y) - 4) < epsilon)
		            	{
							segQuartileCountMatrix[y][3]+= 1;
		            	}
					}
					segementFlag=segmentName==null;
					queueFlag= queueName==null;
				}			
		}
		
		public static void quartilePerformanceKpiAnalysis(MyAgent agent)
		{
			float epsilon=(float) 0.00000001;
			
			boolean segementFlag= segmentName==null;
			
			boolean queueFlag= queueName==null;

				for(int j=0;j<12;j++)
				{					
					if(!segementFlag && j<agent.getSegMonths().size() && j<agent.getSegNameMonths().size())
					{
						//System.out.print(agent.getSegNameMonths().get(j)+"\t");
						segementFlag = agent.getSegNameMonths().get(j).equalsIgnoreCase(segmentName);
						//System.out.println(agent.getSegNameMonths().get(j)+"\t"+segmentName+segementFlag);
					}
					
					if(!queueFlag && j<agent.getSegMonths().size() && j<agent.getQueueNameMonths().size())
					{
						queueFlag = agent.getQueueNameMonths().get(j).equalsIgnoreCase(queueName);
					}
					
					if(j<agent.getSegMonths().size() && segementFlag && queueFlag)
					{
						//System.out.println(agent.getSegNameMonths().get(j)+"\t"+segmentName+segementFlag);
						//ABT
						if(Math.abs(agent.getSegMonths().get(j) - 1) < epsilon && agent.getAbtMonths().get(j)>0)
		            	{
							abtSumMatrix[j][0]+= agent.getAbtMonths().get(j);
							abtCountMatrix[j][0]+= 1;
		            	}
						if(Math.abs(agent.getSegMonths().get(j) - 2) < epsilon && agent.getAbtMonths().get(j)>0)
		            	{
							abtSumMatrix[j][1]+= agent.getAbtMonths().get(j);
							abtCountMatrix[j][1]+= 1;
		            	}
						if(Math.abs(agent.getSegMonths().get(j) - 3) < epsilon && agent.getAbtMonths().get(j)>0)
		            	{
							abtSumMatrix[j][2]+= agent.getAbtMonths().get(j);
							abtCountMatrix[j][2]+= 1;
		            	}
						if(Math.abs(agent.getSegMonths().get(j) - 4) < epsilon && agent.getAbtMonths().get(j)>0)
		            	{
							abtSumMatrix[j][3]+= agent.getAbtMonths().get(j);
							abtCountMatrix[j][3]+= 1;
		            	}	
						
						//FCR
						if(Math.abs(agent.getSegMonths().get(j) - 1) < epsilon && agent.getFcrMonths().get(j)>0)
		            	{
							fcrSumMatrix[j][0]+= agent.getFcrMonths().get(j);
							fcrCountMatrix[j][0]+= 1;
		            	}
						if(Math.abs(agent.getSegMonths().get(j) - 2) < epsilon && agent.getFcrMonths().get(j)>0)
		            	{
							fcrSumMatrix[j][1]+= agent.getFcrMonths().get(j);
							fcrCountMatrix[j][1]+= 1;
		            	}
						if(Math.abs(agent.getSegMonths().get(j) - 3) < epsilon && agent.getFcrMonths().get(j)>0)
		            	{
							fcrSumMatrix[j][2]+= agent.getFcrMonths().get(j);
							fcrCountMatrix[j][2]+= 1;
		            	}
						if(Math.abs(agent.getSegMonths().get(j) - 4) < epsilon && agent.getFcrMonths().get(j)>0)
		            	{
							fcrSumMatrix[j][3]+= agent.getFcrMonths().get(j);
							fcrCountMatrix[j][3]+= 1;
		            	}	
						
						//SV60
						if(Math.abs(agent.getSegMonths().get(j) - 1) < epsilon && agent.getSv60Months().get(j)>0)
		            	{
							sv60SumMatrix[j][0]+= agent.getSv60Months().get(j);
							sv60CountMatrix[j][0]+= 1;
		            	}
						if(Math.abs(agent.getSegMonths().get(j) - 2) < epsilon && agent.getSv60Months().get(j)>0)
		            	{
							sv60SumMatrix[j][1]+= agent.getSv60Months().get(j);
							sv60CountMatrix[j][1]+= 1;
		            	}
						if(Math.abs(agent.getSegMonths().get(j) - 3) < epsilon && agent.getSv60Months().get(j)>0)
		            	{
							sv60SumMatrix[j][2]+= agent.getSv60Months().get(j);
							sv60CountMatrix[j][2]+= 1;
		            	}
						if(Math.abs(agent.getSegMonths().get(j) - 4) < epsilon && agent.getSv60Months().get(j)>0)
		            	{
							sv60SumMatrix[j][3]+= agent.getSv60Months().get(j);
							sv60CountMatrix[j][3]+= 1;
		            	}	
					}
					segementFlag=segmentName==null;
					queueFlag = queueName==null;
				}			
		}
		
		public static void display(String result, Operation op) 
		{
				Sheet sheetTemp=null;
				int rowCount=-1;
				switch(op)
				{
					case QuartilePerformance:
						sheetTemp = sheet2;
						rowCount = rowCountSheet2;
						break;
					case QPNormalized:
						sheetTemp = sheet3;
						rowCount = rowCountSheet3;
						break;												 
					case Attrition:
						sheetTemp = sheet4;
						rowCount = rowCountSheet4;
						break;
				}
			    Row row = sheetTemp.createRow(rowCount);
			    
			    System.out.println(rowCount+"Count"+op);			 
				if(result==null)
				{
					result="PASS";
				}
				row.createCell(0).setCellValue(result);
				row.createCell(6).setCellValue("FAIL");
				row.createCell(12).setCellValue(result);
				row.createCell(18).setCellValue("FAIL");
				row.createCell(24).setCellValue("Difference PASS/FAIL");
				rowCount++;
				
			    row = sheetTemp.createRow(rowCount);
			    int printIndex=0,r=0;
			    while(printIndex<5)
			    {
			    	row.createCell(r+1).setCellValue("Q1");
		    		row.createCell(r+2).setCellValue("Q2");
		    		row.createCell(r+3).setCellValue("Q3");
		    		row.createCell(r+4).setCellValue("Q4");
		    		r+=6;
		    		printIndex++;
			    }
			    
				rowCount++;
				System.out.println(result);
			    for(int y=0;y<segQuartileCountMatrix.length;y++)
	            {
			    	int sum = 0;
			    	int sumFail = 0;
			    	int colCount=0;
			    	int startColIndex=-1;
			    	int endColIndex=-1;
			    	row = sheetTemp.createRow(rowCount);
			    	row.createCell(colCount++).setCellValue("Month-"+(y+1));
	            	for(int z=0;z<segQuartileCountMatrix[0].length;z++)
	            	{
	            		row.createCell(colCount++).setCellValue(segQuartileCountMatrix[y][z]);
	            		System.out.print(segQuartileCountMatrix[y][z]+"\t");
	            		sum+=segQuartileCountMatrix[y][z];
	            	}
	            	colCount++;
	            	row.createCell(colCount++).setCellValue("Month-"+(y+1));
	            	for(int z=0;z<segQuartileFailCountMatrix[0].length;z++)
	            	{	            	
	            		row.createCell(colCount++).setCellValue(segQuartileFailCountMatrix[y][z]);
	            		sumFail+=segQuartileFailCountMatrix[y][z];
	            		//System.out.print(segQuartileFailCountMatrix[y][z]+"\t");
	            	}
	            	colCount++;
	            	row.createCell(colCount++).setCellValue("Month-"+(y+1));
	            	startColIndex=colCount;
	            	for(int z=0;z<segQuartileCountMatrix[0].length;z++)
	            	{	            		            		
	            		if(segQuartileCountMatrix[y][z]==0 && sum==0)
	            			row.createCell(colCount++).setCellValue(0);
	            		else
	            		{
	            			double var1 = Math.round(((double)segQuartileCountMatrix[y][z]/sum)*100 * 100.0) / 100.0;
	            			row.createCell(colCount++).setCellValue(var1+"%");
	            		}	            			
	            		
	            	}

	            	colCount++;
	            	row.createCell(colCount++).setCellValue("Month-"+(y+1));
	            	endColIndex=colCount;
	            	for(int z=0;z<segQuartileFailCountMatrix[0].length;z++)
	            	{	            		            		
	            		if(segQuartileFailCountMatrix[y][z]==0 && sumFail==0)
	            			row.createCell(colCount++).setCellValue(0);
	            		else
	            		{
	            			double var1 = Math.round(((double)segQuartileFailCountMatrix[y][z]/sumFail)*100 * 100.0) / 100.0;
	            			row.createCell(colCount++).setCellValue(var1+"%");
	            		}	            			
	            		
	            	}
	            	colCount++;
	            	row.createCell(colCount++).setCellValue("Month-"+(y+1));
	            	for(int z=0;z<segQuartileFailCountMatrix[0].length;z++)
	            	{	       
	            		double diffPercent=Math.round((Double.parseDouble(row.getCell(startColIndex++).toString().replace("%", "")) - Double.parseDouble(row.getCell(endColIndex++).toString().replace("%", "")))*100.0) / 100.0;
	            		row.createCell(colCount++).setCellValue(diffPercent+"%");
	            		
	            	}	            	            	
	            	rowCount++;
	            	System.out.println();
	            }
			    rowCount++;
		            
	            System.out.println();
	            System.out.println("FAIL");  
	            //row = sheetTemp.createRow(rowCount);
	           // row.createCell(0).setCellValue("FAIL");
	            //rowCount++;
	            for(int y=0;y<segQuartileFailCountMatrix.length;y++)
	            {
	            	//row = sheetTemp.createRow(rowCount);
	            	for(int z=0;z<segQuartileFailCountMatrix[0].length;z++)
	            	{	            	
	            		//row.createCell(z).setCellValue(segQuartileFailCountMatrix[y][z]);
	            		System.out.print(segQuartileFailCountMatrix[y][z]+"\t");
	            	}
	            	//rowCount++;
	            	System.out.println();
	            }
	            //rowCount++;
	            switch(op)
				{
					case QuartilePerformance:
						sheet2 = sheetTemp;
						rowCountSheet2 = rowCount;
						break;
					case QPNormalized:
						sheet3 = sheetTemp;
						rowCountSheet3 = rowCount;
						break;												 
					case Attrition:
						sheet4 = sheetTemp;
						rowCountSheet4 = rowCount;
						break;
				}
		}
		
		public static void displayKpiAnalysis()
		{
			   Row row = sheet.createRow(rowCount);
			   row.createCell(0).setCellValue("ABT");
			   row.createCell(6).setCellValue("ABT %");
			   rowCount++;
			   
			   row = sheet.createRow(rowCount);
			   int printIndex=0,q=0;
			    while(printIndex<2)
			    {
			    	row.createCell(q+1).setCellValue("Q1");
		    		row.createCell(q+2).setCellValue("Q2");
		    		row.createCell(q+3).setCellValue("Q3");
		    		row.createCell(q+4).setCellValue("Q4");
		    		q+=6;
		    		printIndex++;
			    }
			   rowCount++;
			   
			   
			  //ABT		
		      System.out.println("ABT Analysis ");
		      
		    
	            for(int t=0;t<abtSumMatrix.length;t++)
	            {	     
	            	double sum = 0;
			    	double sumFail = 0;
			    	int colCount=0;
			    	int startColIndex=-1;
			    	int endColIndex=-1;
	            	row = sheet.createRow(rowCount);
            	    row.createCell(colCount++).setCellValue("Month-"+(t+1));
            	    
	            	for(int r=0;r<abtSumMatrix[0].length;r++)
	            	{
	            		System.out.print(abtSumMatrix[t][r]/abtCountMatrix[t][r]+"\t");
	            		row.createCell(colCount++).setCellValue(abtSumMatrix[t][r]/abtCountMatrix[t][r]); 
	            		sum+=abtSumMatrix[t][r]/abtCountMatrix[t][r];
	            	}
	    	            	
	            	System.out.println();
	            	colCount++;
	            	row.createCell(colCount++).setCellValue("Month-"+(t+1));
	            	startColIndex=colCount;
	            	for(int z=0;z<abtSumMatrix[0].length;z++)
	            	{	            		            		
	            		if(abtSumMatrix[t][z]==0 && sum==0)
	            			row.createCell(colCount++).setCellValue(0);
	            		else
	            		{
	            			double var1 = Math.round(((double)(abtSumMatrix[t][z]/abtCountMatrix[t][z])/sum) *100 * 100.0) / 100.0;
	            			row.createCell(colCount++).setCellValue(var1+"%");
	            		}	            			
	            		
	            	}            	  
	            	rowCount++;
	            }
	         
	            row = sheet.createRow(rowCount);
	            rowCount++;
	           	            
	               row = sheet.createRow(rowCount);
	               row.createCell(0).setCellValue("FCR "+ result);
				   row.createCell(6).setCellValue("FCR FAIL");
				   row.createCell(12).setCellValue("FCR DIFF");
				   row.createCell(18).setCellValue("FCR "+ result);
					row.createCell(24).setCellValue("FCR FAIL");
					row.createCell(30).setCellValue("Difference PASS/FAIL FCR");
				   rowCount++;
				   
				   row = sheet.createRow(rowCount);
				   printIndex=0;q=0;
				    while(printIndex<6)
				    {
				    	row.createCell(q+1).setCellValue("Q1");
			    		row.createCell(q+2).setCellValue("Q2");
			    		row.createCell(q+3).setCellValue("Q3");
			    		row.createCell(q+4).setCellValue("Q4");
			    		q+=6;
			    		printIndex++;
				    }
				   rowCount++;
				   
	            //FCR	          
	            System.out.println("FCR Analysis "+ result);
	            for(int t=0;t<fcrSumMatrix.length;t++)
	            {
	            	double sum = 0;
	            	double sumFail = 0;
			    	int colCount=0;
			    	int startColIndex=-1;
			    	int endColIndex=-1;
	            	row = sheet.createRow(rowCount);
            	    row.createCell(colCount++).setCellValue("Month-"+(t+1));
	         
	            	for(int r=0;r<fcrSumMatrix[0].length;r++)
	            	{	            		
	            		System.out.print(fcrSumMatrix[t][r]/fcrCountMatrix[t][r]+"\t");
	            		row.createCell(colCount++).setCellValue(fcrSumMatrix[t][r]/fcrCountMatrix[t][r]);
	            		sum+=fcrSumMatrix[t][r]/fcrCountMatrix[t][r];
	            	}	            	
	            	colCount++;
	            	row.createCell(colCount++).setCellValue("Month-"+(t+1));	            	
	            	
	            	for(int r=0;r<fcrSumFailMatrix[0].length;r++)
	            	{
	            		row.createCell(colCount++).setCellValue(fcrSumFailMatrix[t][r]/fcrCountFailMatrix[t][r]);    
	            		sumFail+=fcrSumFailMatrix[t][r]/fcrCountFailMatrix[t][r];
	            	}
	            	colCount++;
	            	row.createCell(colCount++).setCellValue("Month-"+(t+1));
	            	
	            	
	            	for(int r=0;r<fcrSumFailMatrix[0].length;r++)
	            	{
	            		row.createCell(colCount++).setCellValue((fcrSumMatrix[t][r]/fcrCountMatrix[t][r]) - (fcrSumFailMatrix[t][r]/fcrCountFailMatrix[t][r]));	            				            		
	            	}
	            	
	            	colCount++;
	            	row.createCell(colCount++).setCellValue("Month-"+(t+1));
	            	startColIndex=colCount;
	            	for(int z=0;z<fcrSumMatrix[0].length;z++)
	            	{	            		            		
	            		if(fcrSumMatrix[t][z]==0 && sum==0)
	            			row.createCell(colCount++).setCellValue(0);
	            		else
	            		{
	            			double var1 = Math.round(((double)(fcrSumMatrix[t][z]/fcrCountMatrix[t][z])/sum) *100 * 100.0) / 100.0;
	            			row.createCell(colCount++).setCellValue(var1+"%");
	            		}	            			
	            		
	            	}
	            	colCount++;
	            	row.createCell(colCount++).setCellValue("Month-"+(t+1));
	            	endColIndex=colCount;
	            	for(int z=0;z<fcrSumFailMatrix[0].length;z++)
	            	{	            		            		
	            		if(fcrSumFailMatrix[t][z]==0 && sumFail==0)
	            			row.createCell(colCount++).setCellValue(0);
	            		else
	            		{
	            			double var1 = Math.round(((double)(fcrSumFailMatrix[t][z]/fcrCountFailMatrix[t][z])/sumFail)*100 * 100.0) / 100.0;
	            			row.createCell(colCount++).setCellValue(var1+"%");
	            		}	            			
	            		
	            	}
	            	colCount++;
	            	row.createCell(colCount++).setCellValue("Month-"+(t+1));
	            	for(int z=0;z<fcrCountFailMatrix[0].length;z++)
	            	{	       
	            		double diffPercent=Math.round((Double.parseDouble(row.getCell(startColIndex++).toString().replace("%", "")) - Double.parseDouble(row.getCell(endColIndex++).toString().replace("%", "")))*100.0) / 100.0;
	            		row.createCell(colCount++).setCellValue(diffPercent+"%");
	            		
	            	}	            	
	            	
	            	rowCount++;
	            	System.out.println();
	            }
	            row = sheet.createRow(rowCount);
	            rowCount++;
	            
	            System.out.println("FCR Analysis FAIL");
	            for(int t=0;t<fcrSumFailMatrix.length;t++)
	            {
	            	for(int r=0;r<fcrSumFailMatrix[0].length;r++)
	            	{
	            		System.out.print(fcrSumFailMatrix[t][r]/fcrCountFailMatrix[t][r]+"\t");//+ " sum "+abtSumFailMatrix[t][r]+"\t" + " count "+abtCountFailMatrix[t][r]+"\t"
	            	}
	            	System.out.println();
	            }
	            
	            System.out.println("FCR Analysis DIFF");
	            for(int t=0;t<fcrSumFailMatrix.length;t++)
	            {
	            	for(int r=0;r<fcrSumFailMatrix[0].length;r++)
	            	{	            		
	            		System.out.print((fcrSumMatrix[t][r]/fcrCountMatrix[t][r]) - (fcrSumFailMatrix[t][r]/fcrCountFailMatrix[t][r]) + "\t");//+ " sum "+abtSumFailMatrix[t][r]+"\t" + " count "+abtCountFailMatrix[t][r]+"\t"
	            	}
	            	System.out.println();
	            }
	            
	            //SV60
	               row = sheet.createRow(rowCount);
	               row.createCell(0).setCellValue("SV60 "+ result);
				   row.createCell(6).setCellValue("SV60 FAIL");
				   row.createCell(12).setCellValue("SV60 DIFF");
				   row.createCell(18).setCellValue("SV60 "+ result);
					row.createCell(24).setCellValue("SV60 FAIL");
					row.createCell(30).setCellValue("Difference PASS/FAIL SV60");
				   rowCount++;
				   
				   row = sheet.createRow(rowCount);
				   printIndex=0;q=0;
				    while(printIndex<6)
				    {
				    	row.createCell(q+1).setCellValue("Q1");
			    		row.createCell(q+2).setCellValue("Q2");
			    		row.createCell(q+3).setCellValue("Q3");
			    		row.createCell(q+4).setCellValue("Q4");
			    		q+=6;
			    		printIndex++;
				    }
				   rowCount++;
	            System.out.println("SV60 Analysis "+ result);
	            for(int t=0;t<sv60SumMatrix.length;t++)
	            {
	            	double sum = 0;
	            	double sumFail = 0;
			    	int colCount=0;
			    	int startColIndex=-1;
			    	int endColIndex=-1;
	            	row = sheet.createRow(rowCount);
            	    row.createCell(colCount++).setCellValue("Month-"+(t+1));
	            	for(int r=0;r<sv60SumMatrix[0].length;r++)
	            	{
	            		row.createCell(colCount++).setCellValue(sv60SumMatrix[t][r]/sv60CountMatrix[t][r]);
	            		System.out.print(sv60SumMatrix[t][r]/sv60CountMatrix[t][r]+"\t");
	            		sum+=sv60SumMatrix[t][r]/sv60CountMatrix[t][r];
	            	}
	            	colCount++;
	            	row.createCell(colCount++).setCellValue("Month-"+(t+1));
	            	
	            	for(int r=0;r<sv60SumFailMatrix[0].length;r++)
	            	{
	            		row.createCell(colCount++).setCellValue(sv60SumFailMatrix[t][r]/sv60CountFailMatrix[t][r]); 
	            		sumFail+=sv60SumFailMatrix[t][r]/sv60CountFailMatrix[t][r];
	            	}
	            	colCount++;
	            	row.createCell(colCount++).setCellValue("Month-"+(t+1));
	            	
	            	
	            	for(int r=0;r<sv60SumFailMatrix[0].length;r++)
	            	{
	            		row.createCell(colCount++).setCellValue((sv60SumMatrix[t][r]/sv60CountMatrix[t][r]) - (sv60SumFailMatrix[t][r]/sv60CountFailMatrix[t][r]));           			            		
	            	}
	            	
	            	colCount++;
	            	row.createCell(colCount++).setCellValue("Month-"+(t+1));
	            	startColIndex=colCount;
	            	for(int z=0;z<sv60SumMatrix[0].length;z++)
	            	{	            		            		
	            		if(sv60SumMatrix[t][z]==0 && sum==0)
	            			row.createCell(colCount++).setCellValue(0);
	            		else
	            		{
	            			double var1 = Math.round(((double)(sv60SumMatrix[t][z]/sv60CountMatrix[t][z])/sum) *100 * 100.0) / 100.0;
	            			row.createCell(colCount++).setCellValue(var1+"%");
	            		}	            			
	            		
	            	}
	            	colCount++;
	            	row.createCell(colCount++).setCellValue("Month-"+(t+1));
	            	endColIndex=colCount;
	            	for(int z=0;z<sv60SumFailMatrix[0].length;z++)
	            	{	            		            		
	            		if(sv60SumFailMatrix[t][z]==0 && sumFail==0)
	            			row.createCell(colCount++).setCellValue(0);
	            		else
	            		{
	            			double var1 = Math.round(((double)(sv60SumFailMatrix[t][z]/sv60CountFailMatrix[t][z])/sumFail)*100 * 100.0) / 100.0;
	            			row.createCell(colCount++).setCellValue(var1+"%");
	            		}	            			
	            		
	            	}
	            	colCount++;
	            	row.createCell(colCount++).setCellValue("Month-"+(t+1));
	            	for(int z=0;z<sv60CountFailMatrix[0].length;z++)
	            	{	       
	            		double diffPercent=Math.round((Double.parseDouble(row.getCell(startColIndex++).toString().replace("%", "")) - Double.parseDouble(row.getCell(endColIndex++).toString().replace("%", "")))*100.0) / 100.0;
	            		row.createCell(colCount++).setCellValue(diffPercent+"%");
	            		
	            	}	            
	            	rowCount++;
	            	System.out.println();
	            }
	            row = sheet.createRow(rowCount);
	            rowCount++;
	            
	            System.out.println("SV60 Analysis FAIL");
	            for(int t=0;t<sv60SumFailMatrix.length;t++)
	            {
	            	for(int r=0;r<sv60SumFailMatrix[0].length;r++)
	            	{
	            		System.out.print(sv60SumFailMatrix[t][r]/sv60CountFailMatrix[t][r]+"\t");//+ " sum "+abtSumFailMatrix[t][r]+"\t" + " count "+abtCountFailMatrix[t][r]+"\t"
	            	}
	            	System.out.println();
	            }
	            
	            System.out.println("SV60 Analysis DIFF");
	            for(int t=0;t<sv60SumFailMatrix.length;t++)
	            {
	            	for(int r=0;r<sv60SumFailMatrix[0].length;r++)
	            	{            		
	            		System.out.print((sv60SumMatrix[t][r]/sv60CountMatrix[t][r]) - (sv60SumFailMatrix[t][r]/sv60CountFailMatrix[t][r]) + "\t");//+ " sum "+abtSumFailMatrix[t][r]+"\t" + " count "+abtCountFailMatrix[t][r]+"\t"
	            	}
	            	System.out.println();
	            }	            
		}		
		
		public static void displayAttrition() 
		{
		    Row row = sheet4.createRow(rowCountSheet4);
		    		 
			row.createCell(0).setCellValue(result);
			row.createCell(6).setCellValue("FAIL");
			row.createCell(12).setCellValue(result);
			row.createCell(18).setCellValue("FAIL");
			row.createCell(24).setCellValue("Difference "+ result+ " / FAIL");
			rowCountSheet4++;
			
		    row = sheet4.createRow(rowCountSheet4);
		    int printIndex=0,r=0;
		    while(printIndex<5)
		    {
		    	row.createCell(r+1).setCellValue("Q1");
	    		row.createCell(r+2).setCellValue("Q2");
	    		row.createCell(r+3).setCellValue("Q3");
	    		row.createCell(r+4).setCellValue("Q4");
	    		row.createCell(r+5).setCellValue("Total");
	    		r+=6;
	    		printIndex++;
		    }
		    
		    rowCountSheet4++;
			double sumAttrition=0.0d;
        	double sumAttritionFail=0.0d;
			System.out.println(result);
		    for(int y=0;y<segQuartileCountMatrix.length;y++)
            {
		    	//int sum = 0;
		    	//int sumFail = 0;
		    	int colCount=0;
		    	int startColIndex=-1;
		    	int endColIndex=-1;
		    	int startColIndexNew=-1;
		    	int endColIndexNew=-1;
		    	row = sheet4.createRow(rowCountSheet4);
		    	row.createCell(colCount++).setCellValue("Month-"+(y+1));
            	for(int z=0;z<segQuartileCountMatrix[0].length;z++)
            	{
            		row.createCell(colCount++).setCellValue(segQuartileCountMatrix[y][z]);
            		System.out.print(segQuartileCountMatrix[y][z]+"\t");
            		//sum+=segQuartileCountMatrix[y][z];
            	}
            	colCount++;
            	row.createCell(colCount++).setCellValue("Month-"+(y+1));
            	for(int z=0;z<segQuartileFailCountMatrix[0].length;z++)
            	{	            	
            		row.createCell(colCount++).setCellValue(segQuartileFailCountMatrix[y][z]);
            		//sumFail+=segQuartileFailCountMatrix[y][z];
            		//System.out.print(segQuartileFailCountMatrix[y][z]+"\t");
            	}
            	colCount++;
            	row.createCell(colCount++).setCellValue("Month-"+(y+1));
            	startColIndex=colCount;
            	startColIndexNew=colCount;
            	double var2=0;
            	Row tempRow = sheet4.getRow(rowCountSheet4-1);      
            	
            	for(int z=0;z<segQuartileCountMatrix[0].length;z++)
            	{	            		            		
            		//if(segQuartileCountMatrix[y][z]==0)
            			//row.createCell(colCount++).setCellValue(0);
            		//else
            		//{          
            		       		   
            			if(y>0)
            			{
            				var2 = Double.parseDouble(tempRow.getCell(startColIndexNew++).toString().replace("%", ""));
            			}           				
            			double var1 = var2 +Math.round(((double)segQuartileCountMatrix[y][z]/totalAgentsByCategory[0][0])*100 * 100.0) / 100.0;//totalAgentsByCategory[y][resultIndex] changed to [0][resultIndex] for getting accurate %
            			var2=0;
            			sumAttrition+=var1;
            			row.createCell(colCount++).setCellValue(df2.format(var1)+"%");
            		//}	            			            		
            	}
            	
            	row.createCell(colCount++).setCellValue(Math.round((sumAttrition) * 100.0) / 100.0+"%");
            	sumAttrition=0;

            	//startColIndexNew=colCount;
            	//colCount++;
            	row.createCell(colCount++).setCellValue("Month-"+(y+1));
            	endColIndex=colCount;
            	endColIndexNew =colCount;
            	for(int z=0;z<segQuartileFailCountMatrix[0].length;z++)
            	{	            		            		
            		//if(segQuartileFailCountMatrix[y][z]==0)
            			//row.createCell(colCount++).setCellValue(0);
            		//else
            		//{
	            		if(y>0)
	        			{
	        				var2 = Double.parseDouble(tempRow.getCell(endColIndexNew++).toString().replace("%", ""));
	        			}  
            			double var1 = var2 +Math.round(((double)segQuartileFailCountMatrix[y][z]/totalAgentsByCategory[0][totalAgentsByCategory[0].length-1])*100 * 100.0) / 100.0;
            			var2=0;
            			sumAttritionFail+=var1;
            			row.createCell(colCount++).setCellValue(df2.format(var1)+"%");
            		//}	            			
            		
            	}
            	row.createCell(colCount++).setCellValue(Math.round((sumAttritionFail) * 100.0) / 100.0+"%");
            	sumAttritionFail=0;
            	//colCount++;
            	row.createCell(colCount++).setCellValue("Month-"+(y+1));
            	for(int z=0;z<segQuartileFailCountMatrix[0].length+1;z++)
            	{	       
            		double diffPercent=Math.round((Double.parseDouble(row.getCell(startColIndex++).toString().replace("%", "")) - Double.parseDouble(row.getCell(endColIndex++).toString().replace("%", "")))*100.0) / 100.0;
            		row.createCell(colCount++).setCellValue(diffPercent+"%");
            		
            	}	            	            	
            	rowCountSheet4++;
            	System.out.println();
            }
		    rowCountSheet4++;
	            
            System.out.println();
            System.out.println("FAIL");  
            for(int y=0;y<segQuartileFailCountMatrix.length;y++)
            {
            	for(int z=0;z<segQuartileFailCountMatrix[0].length;z++)
            	{	            	
            		System.out.print(segQuartileFailCountMatrix[y][z]+"\t");
            	}
            	System.out.println();
            }
		}
		
		public static void calculateTotalVolume(MyAgent agent)
		{
			ArrayList<Integer> temp = new ArrayList<>();
        	temp.add(agent.getJan31Diff());
        	temp.add(agent.getFeb28Diff());
        	temp.add(agent.getMar31Diff());
        	temp.add(agent.getApr30Diff());
        	temp.add(agent.getMay31Diff());
        	temp.add(agent.getJun30Diff());
        	temp.add(agent.getJul31Diff());
        	temp.add(agent.getAug31Diff());
        	temp.add(agent.getSep30Diff());
        	temp.add(agent.getOct31Diff());
        	temp.add(agent.getNov30Diff());
        	temp.add(agent.getDec31Diff());
        	for(int u=0;u<temp.size();u++)
        	{
        		int x = temp.get(u);
        		if(x>=0 && x<=30)
            	{
        			if(agent.getAttritionFlag()==1)
        			{
        				if(x<=agent.getAttritionCount())
        					totalVolume[0]++;       
        			}
        			else
        				totalVolume[0]++;        			
            	}
            	if(x>30 && x<=60)
            	{
            		if(agent.getAttritionFlag()==1)
        			{
        				if(x<=agent.getAttritionCount())
        					totalVolume[1]++;       
        			}
        			else
        				totalVolume[1]++;  
            	}
            	if(x>60 && x<=90)
            	{
            		if(agent.getAttritionFlag()==1)
        			{
        				if(x<=agent.getAttritionCount())
        					totalVolume[2]++;       
        			}
        			else
        				totalVolume[2]++;  
            	}
            	if(x>90 && x<=120)
            	{
            		if(agent.getAttritionFlag()==1)
        			{
        				if(x<=agent.getAttritionCount())
        					totalVolume[3]++;       
        			}
        			else
        				totalVolume[3]++;  
            	}
            	if(x>120 && x<=150)
            	{
            		if(agent.getAttritionFlag()==1)
        			{
        				if(x<=agent.getAttritionCount())
        					totalVolume[4]++;       
        			}
        			else
        				totalVolume[4]++;  
            	}
            	if(x>150 && x<=180)
            	{
            		if(agent.getAttritionFlag()==1)
        			{
        				if(x<=agent.getAttritionCount())
        					totalVolume[5]++;       
        			}
        			else
        				totalVolume[5]++;  
            	}
            	if(x>180 && x<=210)
            	{
            		if(agent.getAttritionFlag()==1)
        			{
        				if(x<=agent.getAttritionCount())
        					totalVolume[6]++;       
        			}
        			else
        				totalVolume[6]++;  
            	}
            	if(x>210 && x<=240)
            	{
            		if(agent.getAttritionFlag()==1)
        			{
        				if(x<=agent.getAttritionCount())
        					totalVolume[7]++;       
        			}
        			else
        				totalVolume[7]++;  
            	}
            	if(x>240 && x<270)
            	{
            		if(agent.getAttritionFlag()==1)
        			{
        				if(x<=agent.getAttritionCount())
        					totalVolume[8]++;       
        			}
        			else
        				totalVolume[8]++;  
            	}
            	if(x>270 && x<=300)
            	{
            		if(agent.getAttritionFlag()==1)
        			{
        				if(x<=agent.getAttritionCount())
        					totalVolume[9]++;       
        			}
        			else
        				totalVolume[9]++;  
            	}
            	if(x>300 && x<=330)
            	{
            		if(agent.getAttritionFlag()==1)
        			{
        				if(x<=agent.getAttritionCount())
        					totalVolume[10]++;       
        			}
        			else
        				totalVolume[10]++;  
            	}
            	if(x>330 && x<=360)
            	{
            		if(agent.getAttritionFlag()==1)
        			{
        				if(x<=agent.getAttritionCount())
        					totalVolume[11]++;       
        			}
        			else
        				totalVolume[11]++;  
            	}
        	}	            	   
		}
		
		public static void displayTotalVolume()
		{
			 Row row = sheet5.createRow(rowCountSheet5);
		     row.createCell(0).setCellValue("Total Volume");
			 rowCountSheet5++;
		     for(int i=0;i<totalVolume.length;i++)
		     {
		    	 row = sheet5.createRow(rowCountSheet5);
		    	 row.createCell(0).setCellValue("Month-"+(i+1));
		    	 row.createCell(1).setCellValue(totalVolume[i]);		    	 
		    	 rowCountSheet5++;
		     }
		}
		
		public static void attritionAnalysis(ArrayList<Float> list, MyAgent agent)
		{			
			ArrayList<Integer> temp = new ArrayList<>();
        	temp.add(agent.getJan31Diff());
        	temp.add(agent.getFeb28Diff());
        	temp.add(agent.getMar31Diff());
        	temp.add(agent.getApr30Diff());
        	temp.add(agent.getMay31Diff());
        	temp.add(agent.getJun30Diff());
        	temp.add(agent.getJul31Diff());
        	temp.add(agent.getAug31Diff());
        	temp.add(agent.getSep30Diff());
        	temp.add(agent.getOct31Diff());
        	temp.add(agent.getNov30Diff());
        	temp.add(agent.getDec31Diff());
       					
			boolean segementFlag= segmentName==null;
			
			boolean queueFlag= queueName==null;
			
			for(int u=0;u<temp.size();u++)
        	{
        		int x = temp.get(u);
        		if(x>=0 && x<=30)
            	{
        			if(agent.getAttritionFlag()==1)
        			{
        				if(x<=agent.getAttritionCount())
        					tempAttritionCalculation(segementFlag,queueFlag,agent,0);   
        			}
        			else
        				tempAttritionCalculation(segementFlag,queueFlag,agent,0);
            	}
            	if(x>30 && x<=60)
            	{
            		if(agent.getAttritionFlag()==1)
        			{
        				if(x<=agent.getAttritionCount())
        					tempAttritionCalculation(segementFlag,queueFlag,agent,1);   
        			}
        			else
        				tempAttritionCalculation(segementFlag,queueFlag,agent,1);
            	}
            	if(x>60 && x<=90)
            	{
            		if(agent.getAttritionFlag()==1)
        			{
        				if(x<=agent.getAttritionCount())
        					tempAttritionCalculation(segementFlag,queueFlag,agent,2);   
        			}
        			else
        				tempAttritionCalculation(segementFlag,queueFlag,agent,2);
            	}
            	if(x>90 && x<=120)
            	{
            		if(agent.getAttritionFlag()==1)
        			{
        				if(x<=agent.getAttritionCount())
        					tempAttritionCalculation(segementFlag,queueFlag,agent,3);   
        			}
        			else
        				tempAttritionCalculation(segementFlag,queueFlag,agent,3);
            	}
            	if(x>120 && x<=150)
            	{
            		if(agent.getAttritionFlag()==1)
        			{
        				if(x<=agent.getAttritionCount())
        					tempAttritionCalculation(segementFlag,queueFlag,agent,4);   
        			}
        			else
        				tempAttritionCalculation(segementFlag,queueFlag,agent,4);
            	}
            	if(x>150 && x<=180)
            	{
            		if(agent.getAttritionFlag()==1)
        			{
        				if(x<=agent.getAttritionCount())
        					tempAttritionCalculation(segementFlag,queueFlag,agent,5);   
        			}
        			else
        				tempAttritionCalculation(segementFlag,queueFlag,agent,5);
            	}
            	if(x>180 && x<=210)
            	{
            		if(agent.getAttritionFlag()==1)
        			{
        				if(x<=agent.getAttritionCount())
        					tempAttritionCalculation(segementFlag,queueFlag,agent,6);   
        			}
        			else
        				tempAttritionCalculation(segementFlag,queueFlag,agent,6);
            	}
            	if(x>210 && x<=240)
            	{
            		if(agent.getAttritionFlag()==1)
        			{
        				if(x<=agent.getAttritionCount())
        					tempAttritionCalculation(segementFlag,queueFlag,agent,7);   
        			}
        			else
        				tempAttritionCalculation(segementFlag,queueFlag,agent,7);
            	}
            	if(x>240 && x<270)
            	{
            		if(agent.getAttritionFlag()==1)
        			{
        				if(x<=agent.getAttritionCount())
        					tempAttritionCalculation(segementFlag,queueFlag,agent,8);   
        			}
        			else
        				tempAttritionCalculation(segementFlag,queueFlag,agent,8);
            	}
            	if(x>270 && x<=300)
            	{
            		if(agent.getAttritionFlag()==1)
        			{
        				if(x<=agent.getAttritionCount())
        					tempAttritionCalculation(segementFlag,queueFlag,agent,9);   
        			}
        			else
        				tempAttritionCalculation(segementFlag,queueFlag,agent,9);
            	}
            	if(x>300 && x<=330)
            	{
            		if(agent.getAttritionFlag()==1)
        			{
        				if(x<=agent.getAttritionCount())
        					tempAttritionCalculation(segementFlag,queueFlag,agent,10);   
        			}
        			else
        				tempAttritionCalculation(segementFlag,queueFlag,agent,10);
            	}
            	if(x>330 && x<=360)
            	{
            		if(agent.getAttritionFlag()==1)
        			{
        				if(x<=agent.getAttritionCount())
        					tempAttritionCalculation(segementFlag,queueFlag,agent,11);   
        			}
        			else
        				tempAttritionCalculation(segementFlag,queueFlag,agent,11);
            	}
            	segementFlag=segmentName==null;
				queueFlag= queueName==null;	
        	}	      							
		}			
		
		public static void tempAttritionCalculation(boolean segementFlag, boolean queueFlag, MyAgent agent, int month)
		{
			int monthForQS=-1;
			//sb.append(agent.getEmpNo()+","+-1+","+resultIndex+'\n');
			if(!segementFlag)
			{						
				monthForQS=month;
				if(month>=agent.getSegNameMonths().size())
				{
					monthForQS=agent.getSegNameMonths().size()-1;
				}
				segementFlag = agent.getSegNameMonths().get(monthForQS).equalsIgnoreCase(segmentName);
			}
			if(!queueFlag)
			{
				monthForQS=month;
				if(month>=agent.getQueueNameMonths().size())
				{
					monthForQS=agent.getQueueNameMonths().size()-1;
				}					
				queueFlag = agent.getQueueNameMonths().get(monthForQS).equalsIgnoreCase(queueName);
			}
			if(segementFlag && queueFlag)
			{
				totalAgentsByCategoryAttrition[month][0]+=1;		
			}        		
		}
		
		public static void attritionWithoutQuartileAnalysis(int monthNo,MyAgent agent)
		{
			int monthForQS=-1;
			if(monthNo>0)
				monthNo-=1;
			
			boolean segementFlag= segmentName==null;
			
			boolean queueFlag= queueName==null;
			
			if(!segementFlag)
			{			
				monthForQS=monthNo;
				if(monthNo>=agent.getSegNameMonths().size())
				{
					monthForQS=agent.getSegNameMonths().size()-1;
				}
				segementFlag = agent.getSegNameMonths().get(monthForQS).equalsIgnoreCase(segmentName);
			}
			if(!queueFlag)
			{
				monthForQS=monthNo;
				if(monthNo>=agent.getQueueNameMonths().size())
				{
					monthForQS=agent.getQueueNameMonths().size()-1;
				}		
				queueFlag = agent.getQueueNameMonths().get(monthForQS).equalsIgnoreCase(queueName);
			}
			if(segementFlag && queueFlag)
			{							
					AgentsAttritionByCategory[monthNo][0]+=1;
			}
			segementFlag=segmentName==null;
			queueFlag= queueName==null;		
		}
		
		public static void displayAttritionAnalysisWithoutQuartiles()
		{
	           System.out.println("Attrition Matrix");
	           for(int u1=0;u1<AgentsAttritionByCategory.length;u1++)
	           {
	        	   for(int u2=0;u2<AgentsAttritionByCategory[0].length;u2++)
	        	   {
	        		   System.out.print(AgentsAttritionByCategory[u1][u2]+"\t");
	        	   }
	        	   System.out.println();
	           }
	           
	          System.out.println("Attrition Matrix Total");
	           for(int u1=0;u1<totalAgentsByCategoryAttrition.length;u1++)
	           {
	        	   for(int u2=0;u2<totalAgentsByCategoryAttrition[0].length;u2++)
	        	   {
	        		   System.out.print(totalAgentsByCategoryAttrition[u1][u2]+"\t");
	        	   }
	        	   System.out.println();
	           }
	           
			Row row = sheet6.createRow(rowCountSheet6);
			
			int colCount=0;
						
		    double cumulativePercentages=0,cumulativePercentagesFail=0;
		    for(int u2=0;u2<AgentsAttritionByCategory[0].length-1;u2++)
            {		    
	    	  switch(u2)
			   {
				   case 0:
					   row.createCell(1).setCellValue("PASS");
					   break;
				   case 1:
					   row.createCell(1).setCellValue("STRONGLY RECOMMENDED");
					   break;
				   case 2:
					   row.createCell(1).setCellValue("RECOMMENDED");
					   break;
				   case 3:
					   row.createCell(1).setCellValue("RECOMMENDED WITH RESERVATION");
					   break;			   
			   }			
			   row.createCell(2).setCellValue("FAIL");
			   row.createCell(3).setCellValue("Difference");
			   rowCountSheet6++;
			   for(int u1=0;u1<AgentsAttritionByCategory.length;u1++)
        	   {
				   row = sheet6.createRow(rowCountSheet6);
				   colCount=0;
			       row.createCell(colCount++).setCellValue("Month-"+(u1+1));			       
        		   cumulativePercentages+=Math.round(((double)AgentsAttritionByCategory[u1][u2]/totalAgentsByCategoryAttrition[0][u2])*100 * 100.0) / 100.0;
        		   row.createCell(colCount++).setCellValue(df2.format(cumulativePercentages)+"%");        		   
        		   cumulativePercentagesFail+=Math.round(((double)AgentsAttritionByCategory[u1][AgentsAttritionByCategory[0].length-1]/totalAgentsByCategoryAttrition[0][AgentsAttritionByCategory[0].length-1])*100 * 100.0) / 100.0;;
        		   row.createCell(colCount++).setCellValue(df2.format(cumulativePercentagesFail)+"%");
        		   row.createCell(colCount++).setCellValue(df2.format(cumulativePercentages-cumulativePercentagesFail)+"%");
        		   rowCountSheet6++;
        	   }
			  
			   rowCountSheet6++;
			   row = sheet6.createRow(rowCountSheet6);
        	   cumulativePercentages=0;
        	   cumulativePercentagesFail=0;
            }	          
		}
		
		public static void initializeStaticVariables()
		{
			rowCount = 0;rowCountSheet2=0;rowCountSheet3=0;rowCountSheet4=0;rowCountSheet5=0;rowCountSheet6=0;
		}
}
