package com.projectname.scripts;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.util.Properties;

import org.testng.annotations.Test;

import com.projectname.utils.TestConstants;

public class Spilt {

	@Test
	public void spilt() throws Exception{
		
		Properties OR=new Properties();
		FileInputStream fi=new FileInputStream("E:\\SeleniumFrameWorkWithJarBased\\FrameWorkJarFile\\config\\OR.Properties");
		OR.load(fi);
		
		String userName=(String) OR.get("UserName");
		String[] parts=userName.split("_");
		String objectType=parts[0];
		String ojbect=parts[1];
		System.out.println("Username in OR-------------->"+objectType+"------------object name------------------------------"+ojbect);
	}
}
