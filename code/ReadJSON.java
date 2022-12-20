package net.codejava;

import java.util.ArrayList;

import org.json.JSONException;
import org.json.JSONObject;

public class ReadJSON {
	public static void main(String[] args) {
		
	}
	public void readingObject(JSONObject jsonObject) {
		
		String newStringName = jsonObject.optString("name").substring(2, jsonObject.optString("name").length() - 2);
		String[] wordsName = newStringName.split("\",\"");
		
		String newStringNumber = jsonObject.optString("number").substring(2, jsonObject.optString("number").length() - 2);
		String[] wordsNumber = newStringNumber.split("\",\"");
		
		String newStringDate = jsonObject.optString("date").substring(2, jsonObject.optString("date").length() - 2);
		String[] wordsDate = newStringDate.split("\",\"");
		
		ArrayList result = new ArrayList();
		for (int i = 0; i < wordsName.length; i++) {
			result.add(wordsName[i]);
			result.add(wordsNumber[i]);
			result.add(wordsDate[i]);
		}
		//System.out.println("Output data: " + result);       
		//System.out.println("number: " + jsonObject.optString("number"));
		//System.out.println("date: " + jsonObject.optString("date"));
	}
}
