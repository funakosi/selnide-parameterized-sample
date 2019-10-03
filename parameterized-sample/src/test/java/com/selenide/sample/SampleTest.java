package com.selenide.sample;

import static com.codeborne.selenide.Selenide.*;
import static org.hamcrest.CoreMatchers.*;
import static org.junit.Assert.*;

import java.util.Collection;

import org.junit.Before;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.junit.runners.Parameterized;
import org.junit.runners.Parameterized.Parameters;

import com.codeborne.selenide.Configuration;
import com.codeborne.selenide.WebDriverRunner;
import com.example.selenide.page.CheckInfoPage;
import com.example.selenide.page.InputPage;

@RunWith(Parameterized.class)
public class SampleTest {

	private String caseID;

	public SampleTest(String input) {
		this.caseID = input;
	}

	@Before
    public void before(){
        Configuration.browser = WebDriverRunner.GECKO;
        System.setProperty("webdriver.gecko.driver","exe/geckodriver.exe");
    }

	// DataTableを読み込む
	private static void importDataTable() {
		try {
			ExcelUtils.setExcelFile("resources/DataTable.xlsx", "Sheet2");
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	@Parameters(name = "{0}")
	public static Collection<Object[]> TestcaseIdList() {
		importDataTable();
		return ExcelUtils.getTestCases();
	}

	@Test
	public void Test() {
		InputPage inputPage = open("http://example.selenium.jp/reserveApp/", InputPage.class);
		inputPage.setReserveYear(ExcelUtils.getCellValue(caseID, "year"));
        inputPage.setReserveMonth(ExcelUtils.getCellValue(caseID, "month"));
        inputPage.setReserveDay(ExcelUtils.getCellValue(caseID, "day"));
        inputPage.setReserveTerm("2");
        inputPage.setBreakfastOn();
        inputPage.setGuestName("東京太郎");
        inputPage.setPlanA(false);
        inputPage.setPlanA(false);
        CheckInfoPage checkPage = inputPage.clickGotoNext();
        assertThat(checkPage.getErrorCheckResult(), is(""));
        assertThat(checkPage.getDaysCount(), is("2"));
        assertThat(checkPage.getHeadcount(), is("1"));
        assertThat(checkPage.getBfOrder(), is("あり"));
        assertThat(checkPage.getGuestName(), is("東京太郎"));
        assertThat(checkPage.getPrice(), is("17750"));
        checkPage.doCommit();
	}
}
