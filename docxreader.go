package docxfun

import (
	"archive/zip"
	"bytes"
	"errors"
	"fmt"
	"io/ioutil"
	"os"
	"regexp"
	"strconv"
	"strings"

	"github.com/revel/revel"
)

//Docx zip struct
type Docx struct {
	zipFileReader *zip.ReadCloser
	zipReader     *zip.Reader
	Files         []*zip.File
	FilesContent  map[string][]byte
	WordsList     []*Words
}

type Words struct {
	Pid        string   `bson:"Pid,omitempty"`
	RawString  string   `bson:"RawString,omitempty"`
	Content    []string `bson:"Content,omitempty"`
	NewString  string   `bson:"NewString,omitempty"`
	IsNonTran  string   `bson:"IsNonTran,omitempty"` // do not need trans patten
	IsField    string   `bson:"IsField,omitempty"`   // including insrt or other flag indicate it is a updatable field, no need translate
	IsHeading  string   `bson:"IsHeading,omitempty"`
	HeadingLev int      `bson:"HeadingLev,omitempty"`
}

//OpenUploaded
func OpenDocxByte(buff []byte) (*Docx, error) {
	reader, err := zip.NewReader(bytes.NewReader(buff), int64(len(buff)))
	if err != nil {
		return nil, err
	}
	wordDoc := Docx{
		zipReader:    reader,
		Files:        reader.File,
		FilesContent: map[string][]byte{},
	}
	for _, f := range wordDoc.Files {
		contents, _ := wordDoc.retrieveFileContents(f.Name)
		wordDoc.FilesContent[f.Name] = contents
	}

	return &wordDoc, nil
}

//OpenDocx open and load all files content
func OpenDocx(path string) (*Docx, error) {
	reader, err := zip.OpenReader(path)
	if err != nil {
		return nil, err
	}

	wordDoc := Docx{
		zipFileReader: reader,
		Files:         reader.File,
		FilesContent:  map[string][]byte{},
	}

	for _, f := range wordDoc.Files {
		contents, _ := wordDoc.retrieveFileContents(f.Name)
		wordDoc.FilesContent[f.Name] = contents
	}

	return &wordDoc, nil
}

//Close is close reader
func (d *Docx) Close() error {
	return d.zipFileReader.Close()
}

//Read all files contents
func (d *Docx) retrieveFileContents(filename string) ([]byte, error) {
	var file *zip.File
	for _, f := range d.Files {
		if f.Name == filename {
			file = f
		}
	}

	if file == nil {
		return []byte{}, errors.New(filename + " file not found")
	}

	reader, err := file.Open()
	if err != nil {
		return []byte{}, err
	}

	return ioutil.ReadAll(reader)
}

//Save files to new docx file
func (d *Docx) Save(fileName string) error {
	// Create a buffer to write our archive to.
	buf, err := d.ReadToBuffer()
	if err != nil {
		return err
	}
	//write to file
	zipfile, err := os.Create(fileName)
	zipfile.Write(buf.Bytes())
	zipfile.Close()
	return nil
}

//
func (d *Docx) ReadToBuffer() (*bytes.Buffer, error) {
	// Create a buffer to write our archive to.
	buf := new(bytes.Buffer)

	// Create a new zip archive.
	w := zip.NewWriter(buf)

	for fName, content := range d.FilesContent {
		f, err := w.Create(fName)
		if err != nil {
			return buf, err
		}
		_, err = f.Write([]byte(content))
		if err != nil {
			return buf, err
		}
	}
	err := w.Close()
	if err != nil {
		return buf, err
	}
	return buf, nil
}

//Do document replacement, default word/document.xml, replace by paragraphy tags
func (d *Docx) DocumentReplace(fileName string, replaceSlice [][]string) error {
	if fileName == "" {
		fileName = "word/document.xml"
	}
	document := d.FilesContent[fileName]
	docStr := string(document)
	for _, pair := range replaceSlice {
		src := pair[0]
		target := pair[1]
		if len(src) > 0 && len(target) > 0 {
			docStr = strings.Replace(docStr, src, target, -1)
		}
	}
	d.FilesContent["word/document.xml"] = []byte(docStr)
	return nil
}

//GenWordsList
func (d *Docx) GenWordsList() (err error) {
	xmlData := string(d.FilesContent["word/document.xml"])
	listP(xmlData, d)
	return nil
}

//get w:t value
func getT(item string, d *Docx) {
	var subStr string
	data := item
	reRun := regexp.MustCompile(`(?U)(<w:r>|<w:r .*>)(.*)(</w:r>)`)
	re := regexp.MustCompile(`(?U)(<w:t>|<w:t .*>)(.*)(</w:t>)`)
	reFld := regexp.MustCompile(`(?U)<w:fldChar w:fldCharType="begin"/>(.*)<w:fldChar w:fldCharType="end"/>`)
	// reRef := regexp.MustCompile(`(?U)<w:r .* w:val="\S*Reference".*<(w:t|w:t .*)>(.*)</w:t></w:r>`)
	reSimFld := regexp.MustCompile(`(?U)<w:fldSimple .*>.*<(w:t|w:t .*)>(.*)</w:t>.*</w:fldSimple>`)
	reExist := regexp.MustCompile(`(?U){{.*}}`)
	w := new(Words)
	w.RawString = data
	content := []string{}

	wrMatch := reRun.FindAllStringSubmatchIndex(data, -1)

	// loop r
	for _, rMatch := range wrMatch {
		rData := data[rMatch[4]:rMatch[5]]
		fldMatch := reFld.FindAllStringSubmatchIndex(rData, -1)
		simFldMatch := reSimFld.FindAllStringSubmatchIndex(rData, -1)

		wtMatch := re.FindAllStringSubmatchIndex(rData, -1)
		for _, match := range wtMatch {
			subStr = rData[match[4]:match[5]]
			for _, f := range fldMatch {
				//check margine
				if f[0] < match[0] && match[1] < f[1] {
					//text within field, mark as field
					subStr = "{{" + rData[match[4]:match[5]] + "}}"
				}
			}

			for _, f := range simFldMatch {
				existMatch := reExist.FindAllStringSubmatchIndex(subStr, -1)
				if len(existMatch) == 0 {
					if f[0] < match[0] && match[1] < f[1] {
						//text within field, mark as field
						subStr = "{{" + rData[match[4]:match[5]] + "}}"
					}
				}

			}
			//check ConerMarker
			if hasConerMark(rData) {
				fmt.Println(rData)
				if !strings.HasPrefix("{{", subStr) {
					subStr = "{{" + subStr + "}}"
				}
			}
			content = append(content, subStr)
		}
	}

	// mark content which do not need translate
	if isNonTran(strings.Join(content, "")) {
		w.IsNonTran = "Y"
	}

	// mark field
	if hasFldSimple(item) {
		w.IsField = "Y"
	}

	// mark heading
	h, hl := isHeading(item)
	if h {
		w.IsHeading = "Y"
		w.HeadingLev = hl
	}

	w.Content = content
	d.WordsList = append(d.WordsList, w)
}

//isNonTran check if the string should skip for translation.
func isNonTran(item string) bool {
	item = strings.TrimSpace(item)
	r := []string{}
	//for abc123, 2342.123123, absd123.123.asdf
	r = append(r, `(?U)^([0-9]+[a-zA-Z]+|[a-zA-Z]+[0-9.]+)+[0-9a-zA-Z]*$`)
	//for a, A, e, E, only one character
	r = append(r, `(?U)^[a-zA-Z]{1}$`)
	// > < " "
	r = append(r, `^ *(&quote;|&lt;|&gt;|&amp;|&apos;) *$`)
	// SD = 0.5 patten
	r = append(r, `(?U)^[A-Z ]+=?[0-9. ]+$`)
	// on special characters
	r = append(r, `^[0-9_.\-\?\;\,\.\~\*\&\^\%\$\#\@\! \(\)=]+$`)
	// file location
	r = append(r, `^[a-zA-Z]:(\\\w+)*([\\]|[.][a-zA-Z]+)?$`)
	// spaces
	r = append(r, `^ *$`)
	// 3±3, +3-3 format
	r = append(r, `^([\x{00B1}\-+/ =]*[0-9]*)*$`)
	//email
	r = append(r, `^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$`)
	for _, rule := range r {
		re := regexp.MustCompile(rule)
		if re.MatchString(item) {
			fmt.Println(rule)
			fmt.Println(item)
			return true
		}
	}
	return false
}

// hasP identify the paragraph
func hasP(data string) bool {
	re := regexp.MustCompile(`(?U)<w:p .*>(.*)</w:p>`)
	result := re.MatchString(data)
	return result
}

// isHeading flag
func isHeading(data string) (bool, int) {
	// re := regexp.MustCompile(`(?U)<w:pStyle w:val="Heading[0-9]?".*/>`)
	re := regexp.MustCompile(`(?U)<w:pStyle w:val="Heading([0-9]*)".*/>`)
	result := re.MatchString(data)
	if result {
		l := re.FindStringSubmatch(data)
		num, err := strconv.Atoi(l[1])
		if err != nil {
			revel.AppLog.Errorf("Convert heading level error: %v", err)
			return false, 0
		}
		return true, num
	}

	return false, 0
}

//hasFldSimple check "fldSimple" which is field in docx, no translate needed
func hasFldSimple(data string) bool {
	re := regexp.MustCompile(`(?U)<w:fldSimple.*>(.*)</w:fldSimple>`)
	result := re.MatchString(data)
	return result
}

func hasConerMark(data string) bool {
	re := regexp.MustCompile(`(?U)<w:position.*`)
	result := re.MatchString(data)
	return result
}

//hasInstrText check "w:instrText" which is field for paragraphy
func hasInstrText(data string) bool {
	re := regexp.MustCompile(`(?U)<w:instrText.*>(.*)</w:instrText>`)
	result := re.MatchString(data)
	return result
}

// isToc identify TOC which skip trans
func isToc(data string) bool {
	re := regexp.MustCompile(`(?U)<w:pStyle w:val="TOC.*"/>`)
	result := re.MatchString(data)
	return result
}

// TableFigures
func isTbFig(data string) bool {
	re := regexp.MustCompile(`(?U)<w:pStyle w:val="TableofFigures"/>`)
	result := re.MatchString(data)
	return result
}

// listP for w:p tag value
func listP(data string, d *Docx) {
	result := []string{}
	re := regexp.MustCompile(`(?U)<w:p .*>(.*)</w:p>`)
	for _, match := range re.FindAllStringSubmatch(string(data), -1) {
		result = append(result, match[1])
	}
	for _, item := range result {
		if hasP(item) {
			listP(item, d)
			continue
		}
		if isToc(item) || isTbFig(item) {
			continue
		}
		// if hasFldSimple(item) {
		// continue
		// }

		getT(item, d)
	}
	return
}

//Replace with new string to the paragraph
func (c *Words) GenerateNewContent(newString string) {
	//only query in content
	//repalce 1st match
	//remove others with ""
	//How?
	newString = strings.Replace(newString, "\r", "", -1)
	newString = strings.Replace(newString, "\n", "", -1)
	newString = strings.Replace(newString, "\b", "", -1)
	newString = strings.TrimSpace(newString)
	// newString = html.EscapeString(newString)
	// xml clean
	// toTerm = strings.Map(printOnly, toTerm)
	var output string
	output = c.RawString

	var re = regexp.MustCompile(`(?U)(<w:t>|<w:t .*>)(.*)</w:t>`)
	var reFld = regexp.MustCompile(`(?U){{(.*)}}`)
	posList := re.FindAllStringSubmatchIndex(output, -1)

	groups := [][]int{}
	subG := []int{}
	// source group
	for i, item := range c.Content {
		if reFld.MatchString(item) {
			if len(subG) > 0 {
				groups = append(groups, subG)
			}
			groups = append(groups, []int{i})
			subG = []int{}
		} else {
			subG = append(subG, i)
		}
		if i == len(c.Content)-1 {
			if len(subG) > 0 {
				groups = append(groups, subG)
			}
		}
	}

	// target group
	var newGroup = []string{}
	newSlice := reFld.FindAllStringSubmatchIndex(newString, -1)
	// newSlice format [[0,11, 2,9], [....],[....]]
	// len(newSlice) > 0
	for i, item := range newSlice {
		//if {{}} is first
		if i == 0 {
			if item[0] != 0 {
				newGroup = append(newGroup, newString[0:item[0]])
			}
		}
		newGroup = append(newGroup, newString[item[2]:item[3]])
		if i < len(newSlice)-1 {
			// if next one is {{}}
			if item[1] != newSlice[i+1][0] {
				newGroup = append(newGroup, newString[item[1]:newSlice[i+1][0]])
			}
		}
		if i == len(newSlice)-1 {
			if item[1] < len(newString) {
				newGroup = append(newGroup, newString[item[1]:len(newString)])
			}
		}
	}

	// replace content based on group
	if len(groups) == 0 || len(groups) != len(newGroup) {
		for i, pos := range posList {
			left := pos[4]
			right := pos[5]
			if i == 0 {
				output = output[0:left] + newString + output[right:len(output)]
			} else {
				newPlist := re.FindAllStringSubmatchIndex(output, -1)
				if len(newPlist) == len(posList) {
					if len(newPlist[i]) >= 6 {
						left := newPlist[i][4]
						right := newPlist[i][5]
						output = output[0:left] + output[right:len(output)]
					} else {
						revel.AppLog.Errorf("New string regrexp error: %s", output)
					}
				} else {
					revel.AppLog.Errorf("Src and Target Length Not Match: %s", output)
				}
			}
		}
	}

	if len(groups) > 0 && len(groups) == len(newGroup) {
		for j, v := range groups {
			for i, l := range v {
				newPlist := re.FindAllStringSubmatchIndex(output, -1)
				left := newPlist[l][4]
				right := newPlist[l][5]
				if i == 0 {
					output = output[0:left] + filterUnicodeSymbol(newGroup[j]) + output[right:len(output)]
				} else {
					output = output[0:left] + output[right:len(output)]
				}
			}
		}
	}

	c.NewString = output
}

// remove specal symbols
func filterUnicodeSymbol(s string) string {
	return strings.Map(func(r rune) rune {
		switch r {
		case 0x000A, 0x000B, 0x000C, 0x000D, 0x0085, 0x2028, 0x2029, 0xFFFD:
			return -1
		default:
			return r
		}
	}, s)
}
