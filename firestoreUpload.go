package main

import (
	"context"
	"errors"
	"fmt"
	"log"
	"math"
	"os"
	"strconv"
	"strings"
	"time"

	"cloud.google.com/go/firestore"
	firebase "firebase.google.com/go"
	auth "firebase.google.com/go/auth"
	"github.com/tealeg/xlsx"
	"google.golang.org/api/option"
)

func readSheetToSliceOfMap(sheet *xlsx.Sheet) (res []map[string]string, err error) {
	headers := make([]string, 0)
	for i, row := range sheet.Rows {
		if i == 0 {
			for _, cell := range row.Cells {
				str, err := cell.FormattedValue()
				if err != nil {
					return nil, err
				}
				headers = append(headers, fmt.Sprintf("%s", str))
			}
			continue
		}
		vals := make(map[string]string)
		if row != nil {
			p := false
			for _, cell := range row.Cells {
				if cell.String() != "" {
					p = true
					break
				}
			}
			if !p {
				continue
			}
			for j, cell := range row.Cells {
				str, err := cell.FormattedValue()
				if err != nil {
					return nil, err
				}
				vals[headers[j]] = fmt.Sprintf("%s", str)
			}
			res = append(res, vals)
		}
	}

	return res, nil
}

func readFromSourceExcel(filename string) (userlines []map[string]string,
	projectlines []map[string]string,
	measurmentlines []map[string]string,
	designationlines []map[string]string,
	measurementrefslines []map[string]string,
	contactlines []map[string]string, err error) {

	var xlFile *xlsx.File

	xlFile, err = xlsx.OpenFile(filename)
	if err != nil {
		return nil, nil, nil, nil, nil, nil, err
	}

	if len(xlFile.Sheets) == 0 {
		return nil, nil, nil, nil, nil, nil, errors.New("This XLSX file contains no sheets")
	}

	userlines, err = readSheetToSliceOfMap(xlFile.Sheets[0])
	if err != nil {
		return nil, nil, nil, nil, nil, nil, err
	}

	projectlines, err = readSheetToSliceOfMap(xlFile.Sheets[1])
	if err != nil {
		return nil, nil, nil, nil, nil, nil, err
	}

	measurmentlines, err = readSheetToSliceOfMap(xlFile.Sheets[3])
	if err != nil {
		return nil, nil, nil, nil, nil, nil, err
	}

	designationlines, err = readSheetToSliceOfMap(xlFile.Sheets[3])
	if err != nil {
		return nil, nil, nil, nil, nil, nil, err
	}

	measurementrefslines, err = readSheetToSliceOfMap(xlFile.Sheets[3])
	if err != nil {
		return nil, nil, nil, nil, nil, nil, err
	}

	contactlines, err = readSheetToSliceOfMap(xlFile.Sheets[2])
	if err != nil {
		return nil, nil, nil, nil, nil, nil, err
	}

	return userlines, projectlines, measurmentlines, designationlines, measurementrefslines, contactlines, nil
}

func fileExists(name string) bool {
	if _, err := os.Stat(name); err != nil {
		if os.IsNotExist(err) {
			return false
		}
	}
	return true
}

func createLofErrorFile() (logFile *os.File) {
	sdir, err := os.Getwd()
	if err != nil {
		panic(fmt.Sprintf("Fatal error in getting start directory: %v\n", err))
	}
	if fileExists(sdir + "/log_errors.txt") {
		logFile, err = os.OpenFile(sdir+"/log_errors.txt", os.O_RDWR|os.O_CREATE|os.O_APPEND, 0666)
		if err != nil {
			log.Fatal(err)
		}
		return logFile
	}
	logFile, err = os.Create(sdir + "/log_errors.txt")
	if err != nil {
		log.Fatal(err)
	}

	return logFile
}

func doLogError(errStr string) {
	logFile := createLofErrorFile()
	defer logFile.Close()
	log.SetOutput(logFile)
	fmt.Printf("%v \n", errStr)
	log.Fatal(errStr)
}

func roundSpecial(value string) interface{} {
	var x interface{}
	x, err := strconv.ParseFloat(value, 64)
	if err != nil {
		x = value
		return x
	}
	return strconv.FormatFloat(math.Round(x.(float64)*100)/100, 'f', 2, 64)
}

func main() {
	var xlsxPath string
	if len(os.Args) < 2 {
		xlsxPath = "upload_sheet.xlsx"
	} else {
		xlsxPath = os.Args[1]
	}
	fmt.Printf("Use %q file as data source \n", xlsxPath)

	userlines := make([]map[string]string, 0, 0)
	projectlines := make([]map[string]string, 0, 0)
	measurementlines := make([]map[string]string, 0, 0)
	designationlines := make([]map[string]string, 0, 0)
	measurementrefslines := make([]map[string]string, 0, 0)
	contactlines := make([]map[string]string, 0, 0)
	userlines, projectlines, measurementlines, designationlines, measurementrefslines, contactlines, err := readFromSourceExcel(xlsxPath)
	if err != nil {
		doLogError(err.Error())
	}

	opt := option.WithCredentialsFile("serviceAccountKey.json")
	ctx := context.Background()

	app, err := firebase.NewApp(ctx, nil, opt)
	if err != nil {
		doLogError(fmt.Sprintf("Error initializing app: '%v'", err))
	}
	firestoreClient, err := app.Firestore(ctx)
	if err != nil {
		doLogError(err.Error())
	}
	defer firestoreClient.Close()

	if len(userlines) != 0 {
		fmt.Printf("Create user records:")
		authClient, err := app.Auth(ctx)
		if err != nil {
			doLogError(fmt.Sprintf("Error getting Auth client: %v\n", err))
		}
		for _, line := range userlines {
			u, err := authClient.GetUserByEmail(ctx, line["identifier"])
			if err != nil {
				if !strings.Contains(err.Error(), "cannot find user from email") {
					doLogError(fmt.Sprintf("Error getting user by email %s: %v\n", line["identifier"], err))
				}
			}
			if u != nil {
				continue
			}
			fmt.Print(".")
			params := (&auth.UserToCreate{}).
				Email(line["identifier"]).
				EmailVerified(false).
				Password("~1234@56%7&8xlongxx#Vsa232fshort").
				Disabled(false)

			UserRecord, err := authClient.CreateUser(ctx, params)
			if err != nil {
				doLogError(fmt.Sprintf("error creating user: %v\n", err))
			}

			_, err = firestoreClient.Collection("users").Doc(UserRecord.UID).Set(ctx, map[string]interface{}{
				"first_name": line["first_name"],
				"last_name":  line["last_name"],
				"role":       line["role"],
			}, firestore.MergeAll)
		}
	}

	var projectID string
	var totalcables int
	prcollname := "project"
	fmt.Println()
	fmt.Print("Add project details:")
	for _, line := range projectlines {
		fmt.Print(".")
		if line["project_id"] != "" {
			projectID = line["project_id"]
			var startdate interface{}
			if line["start_date"] == "" {
				startdate = nil
			} else {
				startdate, _ = time.Parse("01-02-06", line["start_date"])
			}
			var calibrationdate interface{}
			if line["calibration_date"] == "" {
				calibrationdate = nil
			} else {
				calibrationdate, _ = time.Parse("01-02-06", line["calibration_date"])
			}
			var engineersubmittedat interface{}
			if line["engineer_submitted_at"] == "" {
				engineersubmittedat = nil
			} else {
				engineersubmittedat, _ = time.Parse("01-02-06", line["engineer_submitted_at"])
			}
			var fieldstartedat interface{}
			if line["field_started_at"] == "" {
				fieldstartedat = nil
			} else {
				fieldstartedat, _ = time.Parse("01-02-06", line["field_started_at"])
			}
			var fieldsubmittedat interface{}
			if line["field_submitted_at"] == "" {
				fieldsubmittedat = nil
			} else {
				fieldsubmittedat, _ = time.Parse("01-02-06", line["field_submitted_at"])
			}
			area, _ := strconv.Atoi(line["area"])
			totalcables, _ = strconv.Atoi(line["total_cables"])
			averagedeviation, _ := strconv.Atoi(line["average_deviation"])
			status, _ := strconv.Atoi(line["status"])
			_, err = firestoreClient.Collection(prcollname).Doc(projectID).Set(ctx, map[string]interface{}{
				"address_line_1":           line["address_line_1"],
				"address_line_2":           line["address_line_2"],
				"area":                     area,
				"average_deviation":        averagedeviation,
				"benchmark":                line["benchmark"],
				"calibration_date":         calibrationdate,
				"calibration_psi":          line["calibration_psi"],
				"client_name":              line["client_name"],
				"contact_name":             line["contact_name"],
				"contact_phone":            line["contact_phone"],
				"device_calibration_image": line["device_calibration_image"],
				"engineer_id":              line["engineer_id"],
				"engineer_submitted_at":    engineersubmittedat,
				"field_started_at":         fieldstartedat,
				"field_submitted_at":       fieldsubmittedat,
				"field_tech_id":            line["field_tech_id"],
				"floor":                    line["floor"],
				"gauge":                    line["gauge"],
				"general_location":         line["general_location"],
				"map_image":                line["map_image"],
				"name":                     line["name"],
				"number":                   line["number"],
				"project_id":               projectID,
				"pt_specification":         line["pt_specification"],
				"pump":                     line["pump"],
				"ram":                      line["ram"],
				"ram_certification_image": line["ram_certification_image"],
				"sheet":                   line["sheet"],
				"start_date":              startdate,
				"status":                  status,
				"stressing_company_name":  line["stressing_company_name"],
				"stressing_location":      line["stressing_location"],
				"total_cables":            totalcables,
				"weather":                 line["weather"],
				"work_order_number":       line["work_order_number"],
			}, firestore.MergeAll)
			if err != nil {
				doLogError(fmt.Sprintf("Failed adding %v: %v", line, err))
			}
		}
	}

	fmt.Println()
	fmt.Print("Add measurements:")
	k := 0
	for _, line := range measurementlines {
		cableorder, _ := strconv.Atoi(line["is_second_end"])
		if cableorder != 1 {
			fmt.Print(".")
			isDouble, _ := strconv.ParseBool(line["is_double"])
			_, err = firestoreClient.Collection(prcollname).Doc(projectID).Collection("measurements").
				Doc(projectID+"-"+"measurement"+"-"+strconv.Itoa(k+1)).Set(ctx, map[string]interface{}{
				"designation": roundSpecial(line["Set Designation"]),
				"is_double":   isDouble,
				"cable_id":    line["cable_id"],
			}, firestore.MergeAll)
			if err != nil {
				doLogError(fmt.Sprintf("Failed adding %v: %v", line, err))
			}
			k++
		}
	}

	fmt.Println()
	fmt.Print("Add designations:")
	type designationT struct {
		name         string
		toleranceMax float64
		toleranceMin float64
	}
	designationsunique := make(map[designationT]bool, 0)
	k = 0
	for _, line := range designationlines {
		toleranceMax, _ := strconv.ParseFloat(line["tolerance_max"], 64)
		toleranceMin, _ := strconv.ParseFloat(line["tolerance_min"], 64)
		_, is := designationsunique[designationT{line["Set Designation"], 0, 0}]
		if is {
			continue
		}
		fmt.Print(".")

		_, err = firestoreClient.Collection(prcollname).Doc(projectID).Collection("designations").
			Doc(projectID+"-"+"designation"+"-"+strconv.Itoa(k+1)).Set(ctx, map[string]interface{}{
			"name":          roundSpecial(line["Set Designation"]),
			"tolerance_max": toleranceMax,
			"tolerance_min": toleranceMin,
		}, firestore.MergeAll)
		if err != nil {
			doLogError(fmt.Sprintf("Failed adding %v: %v", line, err))
		}
		designationsunique[designationT{line["Set Designation"], 0, 0}] = true
		k++
	}

	fmt.Println()
	fmt.Print("Add measurement-refs:")
	for j, line := range measurementrefslines {
		fmt.Print(".")
		var cableid string
		cableid = line["cable_id"]
		x, _ := strconv.Atoi(line["x"])
		y, _ := strconv.Atoi(line["y"])
		_, err = firestoreClient.Collection(prcollname).Doc(projectID).Collection("measurement-refs").
			Doc(projectID+"-"+"measurement-ref"+"-"+strconv.Itoa(j+1)).Set(ctx, map[string]interface{}{
			"cable_id": cableid,
			"end_id":   line["end_id"],
			"order_id": j,
			"suffix":   line["suffix"],
			"x":        x,
			"y":        y,
		}, firestore.MergeAll)
		if err != nil {
			doLogError(fmt.Sprintf("Failed adding %v: %v", line, err))
		}
	}

	fmt.Println()
	fmt.Print("Add contacts:")
	for j, line := range contactlines {
		fmt.Print(".")
		if line["email"] != "" {
			status, _ := strconv.Atoi(line["status"])
			_, err = firestoreClient.Collection(prcollname).Doc(projectID).Collection("contacts").
				Doc(projectID+"-"+"contact"+"-"+strconv.Itoa(j+1)).Set(ctx, map[string]interface{}{
				"email":      line["email"],
				"name":       line["name"],
				"statusType": status,
			})
			if err != nil {
				doLogError(fmt.Sprintf("Failed adding %v: %v", line, err))
			}
		}
	}

	fmt.Println()
	fmt.Println("Job done!")
	fmt.Println("Press the Enter Key to quit!")
	var input string
	fmt.Scanln(&input)
}
