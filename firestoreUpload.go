package main

import (
	"context"
	"errors"
	"fmt"
	"log"
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
	designationlines []map[string]string,
	measurementrefslines []map[string]string,
	contactlines []map[string]string, err error) {

	var xlFile *xlsx.File

	xlFile, err = xlsx.OpenFile(filename)
	if err != nil {
		return nil, nil, nil, nil, nil, err
	}

	if len(xlFile.Sheets) == 0 {
		return nil, nil, nil, nil, nil, errors.New("This XLSX file contains no sheets")
	}

	userlines, err = readSheetToSliceOfMap(xlFile.Sheets[0])
	if err != nil {
		return nil, nil, nil, nil, nil, err
	}

	projectlines, err = readSheetToSliceOfMap(xlFile.Sheets[1])
	if err != nil {
		return nil, nil, nil, nil, nil, err
	}

	designationlines, err = readSheetToSliceOfMap(xlFile.Sheets[2])
	if err != nil {
		return nil, nil, nil, nil, nil, err
	}

	measurementrefslines, err = readSheetToSliceOfMap(xlFile.Sheets[3])
	if err != nil {
		return nil, nil, nil, nil, nil, err
	}

	contactlines, err = readSheetToSliceOfMap(xlFile.Sheets[4])
	if err != nil {
		return nil, nil, nil, nil, nil, err
	}

	return userlines, projectlines, designationlines, measurementrefslines, contactlines, nil
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
	designationlines := make([]map[string]string, 0, 0)
	measurementrefslines := make([]map[string]string, 0, 0)
	contactlines := make([]map[string]string, 0, 0)
	userlines, projectlines, designationlines, measurementrefslines, contactlines, err := readFromSourceExcel(xlsxPath)
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
	prcollname := "project"
	fmt.Println()
	fmt.Print("Add project details and measurments:")
	for j, line := range projectlines {
		fmt.Print(".")
		if line["project_id"] != "" {
			projectID = line["project_id"]
			date, _ := time.Parse("01/02/2006", line["start_date"])
			_, err = firestoreClient.Collection(prcollname).Doc(projectID).Set(ctx, map[string]interface{}{
				"address_line_1": line["address_line_1"],
				"address_line_2": line["address_line_2"],
				"area":           line["area"],
				"client_name":    line["client_name"],
				"contact_name":   line["contact_name"],
				"contact_phone":  line["contact_phone"],
				"engineer_id":    line["engineer_id"],
				"field_tech_id":  line["field_tech_id"],
				"map_image":      line["map_image"],
				"project_id":     projectID,
				"start_date":     date,
			}, firestore.MergeAll)
			if err != nil {
				doLogError(fmt.Sprintf("Failed adding %v: %v", line, err))
			}
		}

		isDouble, _ := strconv.ParseBool(line["is_double"])
		cableID, _ := strconv.Atoi(line["cable_id"])
		_, err = firestoreClient.Collection(prcollname).Doc(projectID).Collection("measurements").
			Doc(projectID+"-"+"measurement"+"-"+strconv.Itoa(j+1)).Set(ctx, map[string]interface{}{
			"designation": line["designation"],
			"is_double":   isDouble,
			"cable_id":    cableID,
		}, firestore.MergeAll)
		if err != nil {
			doLogError(fmt.Sprintf("Failed adding %v: %v", line, err))
		}
	}

	fmt.Println()
	fmt.Print("Add designations:")
	for j, line := range designationlines {
		fmt.Print(".")
		toleranceMax, _ := strconv.ParseFloat(line["tolerance_max"], 64)
		toleranceMin, _ := strconv.ParseFloat(line["tolerance_min"], 64)
		_, err = firestoreClient.Collection(prcollname).Doc(projectID).Collection("designations").
			Doc(projectID+"-"+"designation"+"-"+strconv.Itoa(j+1)).Set(ctx, map[string]interface{}{
			"name":          line["designation"],
			"tolerance_max": toleranceMax,
			"tolerance_min": toleranceMin,
		}, firestore.MergeAll)
		if err != nil {
			doLogError(fmt.Sprintf("Failed adding %v: %v", line, err))
		}
	}

	fmt.Println()
	fmt.Print("Add measurement-refs:")
	for j, line := range measurementrefslines {
		fmt.Print(".")
		cableID, _ := strconv.Atoi(line["cable_id"])
		x, _ := strconv.Atoi(line["x"])
		y, _ := strconv.Atoi(line["y"])
		_, err = firestoreClient.Collection(prcollname).Doc(projectID).Collection("measurement-refs").
			Doc(projectID+"-"+"measurement-ref"+"-"+strconv.Itoa(j+1)).Set(ctx, map[string]interface{}{
			"cable_id": cableID,
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
				"email":  line["email"],
				"name":   line["name"],
				"status": status,
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
