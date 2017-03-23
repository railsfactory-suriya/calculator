package main

import (
  "errors"
  "flag"
  "fmt"
  "os"
  "strconv"
  "strings"
  "regexp"
  "math"

  "time"

  "github.com/tealeg/xlsx"
)

const (
  version              string = "0.3"
  errorFile            string = "xlsx2csv.err"
  successFile          string = "xlsx2csv.success"
  logFile              string = "xlsx2csv.log"
  tmpFolder            string = "/tmp"
  sampleInputFilePath  string = "/samplePath/sampleFile.xlsx"
  sampleOutputFilePath string = "/tmp/xlsx2csvOutput.csv"

  tmpErrorFilePath   string = tmpFolder + "/" + errorFile
  tmpSuccessFilePath string = tmpFolder + "/" + successFile
  tmpLogFilePath     string = tmpFolder + "/" + logFile
)

var (
  newSession       bool
  currentFileRowNo int64
  readSheetNo      int
)

// Exists reports whether the named file or directory exists.
func fileExists(name string) bool {
  if _, err := os.Stat(name); err != nil {
    if os.IsNotExist(err) {
      return false
    }
  }
  return true
}

func fileDelete(name string) {
  os.Remove(name)
}

func fileCreate(name string) (f *os.File) {
  f, err := os.Create(name)
  if err != nil {
    panic(err)
  }
  return f
}

func fileOpenToWrite(name string) (f *os.File) {
  f, err := os.OpenFile(name, os.O_RDWR, 0644)
  if err != nil {
    panic(err)
  }
  return f
}

func goError(e error) {
  if e != nil {
    panic(e)
  }
}

func deleteTmpFiles() {
  fileDelete(tmpErrorFilePath)
  fileDelete(tmpSuccessFilePath)
}

func currentTime() string {
  t := time.Now()
  return t.UTC().Format("Jan 2, 2006 at 3:04PM (MST)")
}

func writeError(err error) {
  var f *os.File
  if err != nil {
    if !fileExists(tmpErrorFilePath) {
      f = fileCreate(tmpErrorFilePath)
    } else {
      f = fileOpenToWrite(tmpErrorFilePath)
    }
    defer f.Close()
    if currentFileRowNo > 0 {
      f.Write([]byte("Row No: " + strconv.FormatInt(currentFileRowNo, 10)))
    }
    f.Write([]byte(currentTime() + ": " + err.Error() + "\n"))
  }
}

func writeLog(log string) {
  var f *os.File
  if fileExists(tmpLogFilePath) {
    fileDelete(tmpLogFilePath)
  }
  f = fileCreate(tmpLogFilePath)
  f.Close()
  f = fileOpenToWrite(tmpLogFilePath)
  defer f.Close()
  f.Write([]byte(currentTime() + ": " + "reading row: " + log + ".\n"))
}

func writeSuccess(fileName string) {
  var f *os.File
  if !fileExists(tmpSuccessFilePath) {
    f = fileCreate(tmpSuccessFilePath)
  } else {
    f = fileOpenToWrite(tmpSuccessFilePath)
  }
  defer f.Close()
  f.Write([]byte(currentTime() + ": " + fileName + " successfully printed as CSV.\n"))
}

func init() {
  newSession = true
  currentFileRowNo = -1

  if newSession {
    deleteTmpFiles()
    newSession = false
  }
}

func main() {
    p := fmt.Print
    p(time.Now().Format("Mon Jan _2 15:04:05 2006"))
    fmt.Println(" --- Started")

  inputFile := flag.String("inputFile", sampleInputFilePath, "mention a XLSX file to convert along with path.")
  outputFile := flag.String("outputFile", sampleOutputFilePath, "mention a CSV file to convert along with path.")
  sheetName := flag.String("sheetName", "", "mention the Sheet Name for which the CSV file to be converted.")
  startRowNum := flag.String("startRowNum", "", "mention the Starting Row Number from which the CSV has to be converted.")
  flag.Parse()

  if *inputFile == sampleInputFilePath {
    err := errors.New("inputFile is not missing.")
    writeError(err)
    return
  }

  if !fileExists(*inputFile) {
    err := errors.New(*inputFile + " doesn't exists.")
    writeError(err)
    return
  }

  // #TODO: Need to check if the outputFile can be created.
  outputFilePath := *outputFile
  fileDelete(outputFilePath)
  f := fileCreate(outputFilePath)
  f.Close()
  f = fileOpenToWrite(outputFilePath)
  defer f.Close()

  xlFile, err := xlsx.OpenFile(*inputFile)
  if err != nil {
    writeError(err)
  }

  readSheetNo = -1
  if len(strings.TrimSpace(*sheetName)) > 0 {
    parsedSheetName := strings.TrimSpace(*sheetName)
    for i := 0; i < len(xlFile.Sheets); i++ {
      if xlFile.Sheets[i].Name == parsedSheetName {
        readSheetNo = i
        break
      }
    }
    if readSheetNo == -1 {
      err := errors.New("Sheet : '" + parsedSheetName + "' doesn't exists.")
      writeError(err)
      return
    }
  } else {
    readSheetNo = 0
  }

  readSheetRowNo := 0
  if len(strings.TrimSpace(*startRowNum)) > 0 {
    readSheetRowNo, _ = strconv.Atoi(strings.TrimSpace(*startRowNum))
  }

  currentFileRowNo = 0
  currentSheet := xlFile.Sheets[readSheetNo]

    p(time.Now().Format("Mon Jan _2 15:04:05 2006"))
    fmt.Println("Loaded in memory.")
 
  for _, row := range currentSheet.Rows {
    currentFileRowNo++
    // if IsPrime(int(currentFileRowNo)) != false {
    //     p(time.Now().Format("Mon Jan _2 15:04:05 2006"))
    //  fmt.Printf(" %v is a Prime Number\n", currentFileRowNo)
    // }
    if math.Mod(float64(currentFileRowNo),1000)==0.00 {
      fmt.Println(currentFileRowNo)
    } 
    if currentFileRowNo < int64(readSheetRowNo) {
      continue
    }
    // writeLog(strconv.FormatInt(currentFileRowNo, 10))

    var rowData []string

    for _, cell := range row.Cells {
        _ = thirtyifconditions

        r, _ := regexp.Compile("[mdy]+")
        // Comment this for Debugging purposes
      // fmt.Println(cell.Type(), "- ", cell.Value, "->", cell.GetNumberFormat(), "===>", r.MatchString(cell.GetNumberFormat()))
      cellData := ""


      if r.MatchString(cell.GetNumberFormat())  {
        //cellData = cell.FormattedValue()
        // HACK: Sometimes the date is not parsed properly and it returns something like below
        // strconv.ParseFloat: parsing "6/30/2015": invalid syntax
        //if strings.HasPrefix(cellData, "strconv.ParseFloat: parsing ") && strings.HasSuffix(cellData, ": invalid syntax") {
		{
		  newCellData := strings.Split(cellData, "strconv.ParseFloat: parsing \"")
          newCellData = strings.Split(newCellData[1], "\": invalid syntax")
          cellData = strings.Trim(newCellData[0], " ")
        }

      } else {
        cellData = cell.Value
      }

      // Replace " to "" & later add quotes around the string.
      cellData = "\"" + strings.Replace(cellData, "\"", "\"\"", -1) + "\""
      rowData = append(rowData, cellData)
    }

    // f.Write([]byte(strings.Join(rowData, ",") + "\n"))
  }

  // If no problem then it means success!!!!.
  // writeSuccess(*inputFile)
  fmt.Println("I hope its done!!!!")
  return
}

func thirtyifconditions() (int) {
  if (1==1) {
    if (2==2) {
      if (3==3) {
        if (4==4) {
          if (5==5) { 
            if (6==6) { 
              if (7==7) { 
                if (8==8) { 
                  if (9==9) { 
                    if (10==10) { 
                      if (11==11) { 
                        if (12==12) { 
                          if (13==13) { 
                            if (14==14) { 
                              if (15==15) { 
                                if (16==16) { 
                                  if (17==17) { 
                                    if (18==18) { 
                                      if (19==19) { 
                                        if (20==20) { 
                                          if (21==21) { 
                                            if (22==22) { 
                                              if (23==23) { 
                                                if (24==24) { 
                                                  if (25==25) { 
                                                    if (26==26) { 
                                                      if (27==27) { 
                                                        if (28==28) { 
                                                          if (29==29) { 
                                                            if (30==30) {
                                                              return 1
                                                            }
                                                          }
                                                        }
                                                      }
                                                    }
                                                  } 
                                                }
                                              }
                                            }
                                          }
                                        }
                                      }
                                    }
                                  }
                                }
                              }
                            }
                          }
                        }
                      }
                    }
                  }
                }
              }
            }
          }
        }
      }
    }
  }
  return 5
}

func IsPrime(value int) bool {
    for i := 2; i <= int(math.Floor(float64(value) / 2)); i++ {
        if value%i == 0 {
            return false
        }
    }
    return value > 1
}
