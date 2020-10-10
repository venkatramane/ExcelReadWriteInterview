package myRunner;

import org.junit.runner.RunWith;
import cucumber.api.junit.Cucumber;
import cucumber.api.CucumberOptions;

@RunWith(Cucumber.class)
@CucumberOptions(
		features="C:\\Users\\VENKATRAMAN\\Downloads\\LexisNexis\\LexisNexis\\src\\main\\java\\Feature\\excel.feature",
		glue={"stepDefinition"}
		)


public class Runner {

}
