package celfiche

import (
	"fmt"
	"strconv"
	"strings"
	"time"

	"github.com/360EntSecGroup-Skylar/excelize/v2"
	"github.com/mxschmitt/playwright-go"
)

type celfiche struct {
	pw      *playwright.Playwright
	browser playwright.Browser
	page    playwright.Page
}

type formData struct {
	LabelName       string
	VariableName    string
	ClassName       string
	TypeName        string
	MultiLineHeight int
	Options         []string
	Iteration       int
}

func NewClient(url string, headless bool) (*celfiche, error) {
	pw, err := playwright.Run()
	if err != nil {
		return nil, fmt.Errorf("Cannot create new playwright instance %v", err)
	}

	launchOpts := playwright.BrowserTypeLaunchOptions{
		Headless: playwright.Bool(headless),
	}

	browser, err := pw.Chromium.Launch(launchOpts)
	if err != nil {
		return nil, fmt.Errorf("Error encountered on Chromium launch %v", err)
	}

	page, err := browser.NewPage()
	if err != nil {
		return nil, err
	}

	pgOpts := playwright.PageGotoOptions{
		WaitUntil: playwright.String("networkidle"),
	}

	_, err = page.Goto(url, pgOpts)
	if err != nil {
		return nil, err
	}

	var c celfiche

	c.pw = pw
	c.browser = browser
	c.page = page

	return &c, nil

}

func (c celfiche) Login(username string, password string) error {
	err := c.page.Type("#UserName", username)
	if err != nil {
		return err
	}

	err = c.page.Type("#Password", password)
	if err != nil {
		return err
	}

	err = c.page.Click("#submit")
	if err != nil {
		return err
	}

	return nil
}

func validateExcel(e *excelize.File) ([]formData, error) {

	consecutiveBlanks := 0
	cellLine := 2
	finalForm := make([]formData, 0)

	for {
		var f formData

		if consecutiveBlanks >= 2 {
			break
		}

		labelCell := fmt.Sprintf("A%v", cellLine)
		variableCell := fmt.Sprintf("B%v", cellLine)
		classNameCell := fmt.Sprintf("C%v", cellLine)
		typeCell := fmt.Sprintf("D%v", cellLine)
		multiLineHeightCell := fmt.Sprintf("E%v", cellLine)
		optionsCell := fmt.Sprintf("F%v", cellLine)
		iterationCell := fmt.Sprintf("G%v", cellLine)

		labelName, err := e.GetCellValue("Sheet1", labelCell)
		if err != nil {
			return nil, fmt.Errorf("Error occurred on line %v %v", cellLine, err)
		}

		variableName, err := e.GetCellValue("Sheet1", variableCell)
		if err != nil {
			return nil, fmt.Errorf("Error occurred on line %v %v", cellLine, err)
		}

		className, err := e.GetCellValue("Sheet1", classNameCell)
		if err != nil {
			return nil, fmt.Errorf("Error occurred on line %v %v", cellLine, err)
		}

		typeName, err := e.GetCellValue("Sheet1", typeCell)
		if err != nil {
			return nil, fmt.Errorf("Error occurred on line %v %v", cellLine, err)
		}

		multiLineHeight, err := e.GetCellValue("Sheet1", multiLineHeightCell)
		if err != nil {
			return nil, fmt.Errorf("Error occurred on line %v %v", cellLine, err)
		}

		optionsValue, err := e.GetCellValue("Sheet1", optionsCell)
		if err != nil {
			return nil, fmt.Errorf("Error occurred on line %v %v", cellLine, err)
		}

		iterationString, err := e.GetCellValue("Sheet1", iterationCell)
		if err != nil {
			return nil, fmt.Errorf("Error occurred on line %v %v", cellLine, err)
		}

		if len(labelName) == 0 {
			consecutiveBlanks = consecutiveBlanks + 1
			continue
		}

		if len(typeName) == 0 {
			return nil, fmt.Errorf("Error occurred on line %v. Type name cannot be blank", cellLine)
		}

		if typeName == "Multi-line" {

			f.MultiLineHeight = 3

			if len(multiLineHeight) > 0 {
				multiLineInt, err := strconv.Atoi(multiLineHeight)
				if err != nil {
					return nil, fmt.Errorf("Error occurred on line %v, Multi line height must be an int", cellLine)
				}

				f.MultiLineHeight = multiLineInt
			}

		}

		if typeName == "Radio Button" || typeName == "Checkbox" || typeName == "Drop-down" {

			if len(optionsValue) > 0 {
				optionsValueSplit := strings.Split(optionsValue, "|")

				for _, v := range optionsValueSplit {
					f.Options = append(f.Options, v)
				}
			}
		}

		f.Iteration = 1

		if len(iterationString) > 0 {
			iterationInt, err := strconv.Atoi(iterationString)
			if err != nil {
				return nil, fmt.Errorf("Error occurred on line %v, %v", cellLine, err)
			}

			if iterationInt <= 0 {
				iterationInt = 1
			}

			f.Iteration = iterationInt
		}

		cellLine = cellLine + 1
		consecutiveBlanks = 0

		f.LabelName = labelName
		f.VariableName = variableName
		f.ClassName = className
		f.TypeName = typeName

		finalForm = append(finalForm, f)
	}

	return finalForm, nil
}

func (c celfiche) ConvertExcel(formURL string, excelPath string, pauseSleep int) error {
	pgOpts := playwright.PageGotoOptions{
		WaitUntil: playwright.String("networkidle"),
	}

	_, err := c.page.Goto(formURL, pgOpts)
	if err != nil {
		return fmt.Errorf("Error navigating to %v %v", formURL, err)
	}

	excelHandler, err := excelize.OpenFile(excelPath)
	if err != nil {
		return fmt.Errorf("Error encountered while open excel file %v", err)
	}

	form, err := validateExcel(excelHandler)
	if err != nil {
		return err
	}

	for _, v := range form {
		for i := 1; i <= v.Iteration; i++ {
			iterString := strconv.Itoa(i)
			clickElement := fmt.Sprintf(`[title="%v"]`, v.TypeName)
			c.page.Click(clickElement)
			c.page.Type("#fieldLabelInput", strings.ReplaceAll(v.LabelName, `${i}`, iterString))
			c.page.Type("#memberName", strings.ReplaceAll(v.VariableName, `${i}`, iterString))

			if v.TypeName == "Multi-line" {
				heightString := strconv.Itoa(v.MultiLineHeight)
				c.page.Press(`[ng-model="fem.currentField.rows"]`, "Backspace")
				c.page.Type(`[ng-model="fem.currentField.rows"]`, heightString)
			}

			if len(v.Options) > 0 {
				for i := 0; i <= 50; i++ {
					c.page.Press("#autoSuggestions", "Backspace")
				}

				for _, o := range v.Options {
					c.page.Type("#autoSuggestions", o)
					c.page.Press("#autoSuggestions", "Enter")
				}

				c.page.Press("#autoSuggestions", "Enter")
			}

			if len(v.ClassName) > 0 {
				c.page.Click(`[ng-click="fem.tab = 2"]`)
				c.page.Type(`[ng-model="fem.currentField.classNames"]`, strings.ReplaceAll(v.ClassName, `${i}`, iterString))
			}

			c.page.Click(".fcClose")

			if pauseSleep > 0 {
				time.Sleep(time.Duration(pauseSleep) * time.Second)
			}
		}
	}

	return nil
}

func (c celfiche) Stop() error {
	err := c.browser.Close()
	if err != nil {
		return fmt.Errorf("There was a problem shutting down celfiche %v", err)
	}

	err = c.pw.Stop()
	if err != nil {
		return fmt.Errorf("There was a problem shutting down celfiche %v", err)
	}

	return nil
}
