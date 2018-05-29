package cohort.analysis;

import java.util.Random;
import java.text.DateFormatSymbols;
public class RandomGenerator {
	public static void main(String args[])
	{
		Random ran = new Random();
		int months[] = new int[40];
		
		for(int i=0;i<40;i++)
		{
			int noOfMonths = ran.nextInt(12) + 1;
			
			if(i<20)
				noOfMonths=12;
			if(i<31 && i>20)
				noOfMonths=ran.nextInt(3) + 1;
			months[i] = noOfMonths;
			for(int j=0;j<noOfMonths;j++)
			{
				System.out.print(ran.nextInt(4) + 1+"\t");
			}
			System.out.println();
		}
		
		for(int i=0;i<40;i++)
		{
			int month=0;
			if(12-months[i]!=0)
			 month = ran.nextInt(12-months[i]) + 1;
			System.out.print(ran.nextInt(28) + 1+"-"+getMonth(month).substring(0,3)+"-17"+"\t");
			if((month+months[i])/12>0)
			{
				if(months[i]<11)
				System.out.print(ran.nextInt(28) + 1+"-"+getMonth((months[i]+month)%12).substring(0,3)+"-18");
			}				
			else
				System.out.print(ran.nextInt(28) + 1+"-"+getMonth(months[i]+month-1).substring(0,3)+"-17");
			//
			//
			System.out.println();
			
		}
	}
	public static String getMonth(int month) {
	    return new DateFormatSymbols().getMonths()[month];
	}
}
