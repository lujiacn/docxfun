package docxfun

import (
	"fmt"
	"testing"
)

func TestRead(t *testing.T) {
	doc, err := OpenDocx("test/test.docx")
	if err != nil {
		fmt.Println("err during reading:", err)
	}

	doc.Close()
	// xmlData := doc.filesContent["word/document.xml"]
	// mv, err := mxj.NewMapXml(xmlData)
	// if err != nil {
	// fmt.Println("Parse xml error: ", err)
	// }

	// tSlice, err := mv.ValuesForKey("t")
	// if err != nil {
	// fmt.Println("Search key error: ", err)
	// }
	// for _, v := range tSlice {
	// fmt.Println("text is ", v)
	// }
	// replaceMap := map[string]string{"hello": "你好", "good": "很好", "<test></test>": "test replacement"}
	// err = doc.DocumentReplace("word/document.xml", replaceMap)
	// if err != nil {
	// fmt.Println("Replace err", err)
	// }
	// err = doc.Save("test/test2.docx")
	// if err != nil {
	// fmt.Println("save err", err)
	// }
	//--------
	err = doc.GetWording()
	if err != nil {
		fmt.Println("list word err", err)
	}
	for _, item := range doc.WordsList {
		fmt.Println(item.Pid, len(item.Content), item.Content)
	}

}
