package main;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.Map;

public class sectorMapping {
	public Map<Integer, String> kMap=new HashMap<Integer,String>();
	
	public sectorMapping(){
		
	}
	
	public Map<Integer,String> uniqSectorMapping(){
		sector ss=new sector();
		ArrayList<String> sectors=ss.sectorValues();
		int ii=0;
		for(int i=1;i<sectors.size();i++){
			
			if(!this.kMap.containsValue(sectors.get(i))){
				this.kMap.put(ii, sectors.get(i));
				ii++;
			}
			
		}
		return this.kMap;
	}

}
