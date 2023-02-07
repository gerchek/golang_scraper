package main

import (
	"encoding/json"
	"fmt"
	"io/ioutil"
	"log"
	"net/http"
	"os"
	"strconv"

	"github.com/360EntSecGroup-Skylar/excelize"
)

type Item struct {
	Version uint       `json:"version"`
	Params  Paramsinfo `json:"params"`
	Data    Datainfo   `json:"data"`
}

type Paramsinfo struct {
	Curr string `json:"curr"`
}

type Datainfo struct {
	Products []Productsinfo `json:"products"`
}

type Productsinfo struct {
	Id         int    `json:"id"`
	Name       string `json:"name"`
	SalePriceU int    `json:"salePriceU"`
	Brand      string `json:"brand"`
}

func main() {

	if len(os.Args) == 2 {
		if os.Args[1] == "1" {
			parsingItems()
		}
		if os.Args[1] == "2" {
			parsingItem()
		}
		fmt.Println("Exit...")
	} else {
		fmt.Println("Error")
	}

}

func parsingItems() {
	file := excelize.NewFile()
	file, err := excelize.OpenFile("items.xlsx")
	url := file.GetCellValue("Sheet1", "A1")
	if err != nil {
		log.Fatal(err)
	}
	// fmt.Println(c1)
	var item Item
	resp, err := http.Get(url)
	if err != nil {
		log.Fatalln(err)
	}
	//We Read the response body on the line below.
	body, err := ioutil.ReadAll(resp.Body)
	if err != nil {
		log.Fatalln(err)
	}
	//Convert the body to type string
	err = json.Unmarshal(body, &item)
	if err != nil {
		panic(err)
	}

	for index, element := range item.Data.Products {
		column := strconv.FormatInt(int64(index+5), 10)
		a_column := "A" + column
		b_column := "B" + column
		c_column := "C" + column
		d_column := "D" + column
		file.SetCellValue("Sheet1", a_column, element.Id)
		file.SetCellValue("Sheet1", b_column, element.Name)
		file.SetCellValue("Sheet1", c_column, element.SalePriceU/100)
		file.SetCellValue("Sheet1", d_column, element.Brand)
	}
	if err := file.SaveAs("items.xlsx"); err != nil { //checking for er>
		log.Fatal(err)
	}
}

func parsingItem() {
	file, err := excelize.OpenFile("item.xlsx")
	if err != nil {
		log.Fatal(err)
	}
	index := 2
	for {
		column := strconv.FormatInt(int64(index), 10)
		id := file.GetCellValue("Sheet1", "A"+column)
		if id == "" {
			fmt.Println("Breaking out of loop")
			break // break here
		}
		// --------------------------------------------------fecth
		var item Item
		url := "https://card.wb.ru/cards/detail?spp=0&regions=80,64,83,4,38,33,70,68,69,86,75,30,40,48,1,66,31,22,71&pricemarginCoeff=1.0&reg=0&appType=1&emp=0&locale=ru&lang=ru&curr=rub&couponsGeo=12,3,18,15,21&dest=-1029256,-102269,-2162196,-1257786&nm=" + id
		resp, err := http.Get(url)
		if err != nil {
			log.Fatalln(err)
		}
		//We Read the response body on the line below.
		body, err := ioutil.ReadAll(resp.Body)
		if err != nil {
			log.Fatalln(err)
		}
		//Convert the body to type string
		err = json.Unmarshal(body, &item)
		if err != nil {
			panic(err)
		}
		b_column := "B" + column
		c_column := "C" + column
		d_column := "D" + column
		file.SetCellValue("Sheet1", b_column, item.Data.Products[0].Name)
		file.SetCellValue("Sheet1", c_column, item.Data.Products[0].SalePriceU/100)
		file.SetCellValue("Sheet1", d_column, item.Data.Products[0].Brand)

		index++
	}
	if err := file.SaveAs("item.xlsx"); err != nil { //checking for er>
		log.Fatal(err)
	}
}
