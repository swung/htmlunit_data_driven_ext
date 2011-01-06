package org.htmlunit.ddext;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Collection;
import java.util.List;

import org.junit.Test;
import org.junit.runner.RunWith;
import org.junit.runners.Parameterized;
import org.junit.runners.Parameterized.Parameters;

import com.gargoylesoftware.htmlunit.util.NameValuePair;

@RunWith(Parameterized.class)
public class DataDriverTest
{
	@Sheet
	public List<NameValuePair> createSubmitRequest;

	@Sheet
	public List<NameValuePair> buildConfigRequest;

	@Sheet
	public List<NameValuePair> saveJobConfigRequest;

	public DataDriverTest(List<NameValuePair> createSubmitRequest, List<NameValuePair> buildConfigRequest,
			List<NameValuePair> saveJobConfigRequest)
	{
		this.createSubmitRequest = createSubmitRequest;
		this.buildConfigRequest = buildConfigRequest;
		this.saveJobConfigRequest = saveJobConfigRequest;
	}

	@Parameters
	public static Collection spreadsheetData() throws IOException
	{
		InputStream spreadsheet = new FileInputStream("walle.xls");
		return new DataDriver(spreadsheet, DataDriverTest.class).getData();
	}

	@Test
	public void testParams() throws Exception
	{
		System.out.println(createSubmitRequest.get(0));
		System.out.println(buildConfigRequest.get(0));
		System.out.println(buildConfigRequest.get(2));
		System.out.println(saveJobConfigRequest.get(0));
	}
}
