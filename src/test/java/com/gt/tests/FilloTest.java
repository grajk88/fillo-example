package com.gt.tests;

import org.testng.annotations.AfterSuite;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

import com.aventstack.extentreports.ExtentReports;
import com.aventstack.extentreports.Status;
import com.aventstack.extentreports.reporter.ExtentSparkReporter;
import com.codoid.products.exception.FilloException;
import com.codoid.products.fillo.Connection;
import com.codoid.products.fillo.Fillo;
import com.codoid.products.fillo.Recordset;

public class FilloTest {

	Connection connection;
	Fillo fillo = new Fillo();
	ExtentReports extent = new ExtentReports();
	ExtentSparkReporter spark = new ExtentSparkReporter("reports/Report.html");

	@BeforeTest
	public void beforeTest() {

		try {

			connection = fillo.getConnection(System.getProperty("user.dir") + "//data//Name_Details.xlsx");

			extent.attachReporter(spark);

		} catch (FilloException e) {

			e.printStackTrace();
		}
	}

	@Test(enabled = true)
	public void extractData() {

		try {

			String strQuery = "Select * from Sheet1";
			Recordset recordset = connection.executeQuery(strQuery);

			while (recordset.next()) {
				System.out.println(recordset.getField("Name"));
				System.out.println(recordset.getField("Age"));
				System.out.println(recordset.getField("Sex"));
			}

			recordset.close();

			extent.createTest("Extract Data").log(Status.FAIL, "Yay! Passed");

		} catch (Exception e) {
			e.printStackTrace();
		}

	}

	@Test
	public void writeData() {

		try {

			String strQuery = "Update Sheet1 Set Result='PASS' where name='Vaishnavi'";

			connection.executeUpdate(strQuery);

			extent.createTest("Write Data").log(Status.PASS, "Yay! Passed");

		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	@Test
	public void addData() {

		try {

			String strQuery = "INSERT INTO Sheet1(Name,Age,Sex) VALUES('Mickey','100','T')";

			connection.executeUpdate(strQuery);

			extent.createTest("Insert Data").log(Status.PASS, "Yay! Passed");

		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	@Test
	public void removeData() {

		try {

			String strQuery = "DELETE FROM Sheet1 WHERE Name='Peter'";

			connection.executeUpdate(strQuery);

			extent.createTest("Remove Data").log(Status.PASS, "Yay! Passed");

		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	@AfterSuite
	public void afterSuite() {
		connection.close();

		extent.flush();
	}

}
