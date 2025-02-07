package main

import (
	"encoding/csv"
	"flag"
	"fmt"
	"io"
	"log"
	"os"
	"path/filepath"
	"strconv"
	"strings"
	"time"
	"unicode/utf8"

	"github.com/tealeg/xlsx"
	"golang.org/x/text/encoding/japanese"
	"golang.org/x/text/transform"
	"golang.org/x/text/unicode/norm"
)

func failOnError(err error) {
	if err != nil {
		log.Fatal("Error:", err)
	}
}

func main() {
	flag.Parse()

	// ログファイル準備
	logfile, err := os.OpenFile("./log.txt", os.O_APPEND|os.O_CREATE|os.O_WRONLY, os.ModePerm)
	failOnError(err)
	defer logfile.Close()

	log.SetOutput(logfile)

	log.Print("Start\r\n")

	// ファイルを読み込んで二次元配列に入れる
	filePath := flag.Arg(0)
	records := readfile(filePath)

	// 出力するフォルダを作成
	filePath = dirCreate(filePath)

	// データの変換 健康診断
	dataConversion(filePath, records)

	// データの変換 胃がん
	gastricConversion(filePath, records)

	// データの変換 子宮がん
	uterineConversion(filePath, records)

	// データの変換 乳がん
	breastConversion(filePath, records)

	// データの変換 前立腺がん
	prostateConversion(filePath, records)

	// データの変換 マンモグラフィー
	mmgConversion(filePath, records)

	// データの変換 骨密度
	dexaConversion(filePath, records)

	log.Print("Finish !\r\n")

}

func readfile(filename string) [][]string {
	//入力ファイル準備
	infile, err := os.Open(filename)
	failOnError(err)
	defer infile.Close()

	reader := csv.NewReader(transform.NewReader(infile, japanese.ShiftJIS.NewDecoder()))
	reader.Comma = '\t'

	//CSVファイルを２次元配列に展開
	readrecords := make([][]string, 0)
	for {
		record, err := reader.Read() // 1行読み出す
		if err == io.EOF {
			break
		} else {
			failOnError(err)
		}

		readrecords = append(readrecords, record)
	}

	return readrecords
}

func dirCreate(path string) string {
	day := time.Now()
	outDir, _ := filepath.Split(path)
	outDirPlus := outDir + "/松英会職員健診データ" + day.Format("20060102")

	if err := os.Mkdir(outDirPlus, 0777); err != nil {
		log.Print(outDirPlus + "\r\n")
		log.Print("出力先のディレクトリを作成できませんでした\r\n")
		return outDir
	} else {
		return outDirPlus + "/"
	}
}

func dataConversion(filename string, inRecs [][]string) {
	// var excelFile *xlsx.File
	// var sheet *xlsx.Sheet
	var vcell *xlsx.Cell
	var cell string

	recLen := 98 //出力する項目数
	cRec := make([]string, recLen)
	var I int

	day := time.Now()

	excelName, _ := filepath.Split(filename)
	excelName = excelName + "松英会職員健診データ" + day.Format("20060102") + ".xlsx"
	excelFile := xlsx.NewFile()
	xlsx.SetDefaultFont(11, "游ゴシック")
	sheet, err := excelFile.AddSheet("データ")
	failOnError(err)

	//タイトル行
	cRec[0] = "実施健診機関CD"
	cRec[1] = "健診種別CD"
	cRec[2] = "受診日"
	cRec[3] = "事業所記号"
	cRec[4] = "証番号"
	cRec[5] = "資格区分"
	cRec[6] = "続柄"
	cRec[7] = "枝番"
	cRec[8] = "漢字氏名"
	cRec[9] = "カナ氏名"
	cRec[10] = "性別"
	cRec[11] = "生年月日"
	cRec[12] = "OP　０１"
	cRec[13] = "OP　０２"
	cRec[14] = "OP　０３"
	cRec[15] = "OP　０４"
	cRec[16] = "OP　０５"
	cRec[17] = "OP　０６"
	cRec[18] = "OP　０７"
	cRec[19] = "OP　０８"
	cRec[20] = "OP　０９"
	cRec[21] = "OP　１０"
	cRec[22] = "OP 11"
	cRec[23] = "請求区分"
	cRec[24] = "健診金額"
	cRec[25] = "法定金額"
	cRec[26] = "請求金額"
	cRec[27] = "支払先CD"
	cRec[28] = "身長"
	cRec[29] = "体重"
	cRec[30] = "BMI"
	cRec[31] = "腹囲"
	cRec[32] = "身体検査判定"
	cRec[33] = "血圧（収縮期）"
	cRec[34] = "血圧（拡張期）"
	cRec[35] = "空腹時中性脂肪"
	cRec[36] = "随時中性脂肪"
	cRec[37] = "HDL・CO"
	cRec[38] = "LDL・CO"
	cRec[39] = "Non・HDLCO"
	cRec[40] = "AST(GOT)"
	cRec[41] = "ALT(GPT)"
	cRec[42] = "γ・GTP"
	cRec[43] = "空腹時血糖"
	cRec[44] = "HｂA1ｃ"
	cRec[45] = "随時血糖"
	cRec[46] = "採血時間"
	cRec[47] = "尿糖"
	cRec[48] = "尿蛋白"
	cRec[49] = "未実施の場合その理由"
	cRec[50] = "白血球数"
	cRec[51] = "赤血球数"
	cRec[52] = "血色素量"
	cRec[53] = "ヘマトクリット"
	cRec[54] = "心電図所見"
	cRec[55] = "眼底精密所見"
	cRec[56] = "血清クレアチニン"
	cRec[57] = "eGFR"
	cRec[58] = "HBｓ抗原"
	cRec[59] = "HBs抗体"
	cRec[60] = "HCV抗体価精密測定"
	cRec[61] = "胸部X線検査判定"
	cRec[62] = "尿酸値"
	cRec[63] = "腹部超音波検査判定"
	cRec[64] = "便潜血"
	cRec[65] = "総合判定"
	cRec[66] = "メタボリック判定"
	cRec[67] = "医師の診断"
	cRec[68] = "医師名"
	cRec[69] = "既往歴"
	cRec[70] = "具体的な既往歴"
	cRec[71] = "自覚症状"
	cRec[72] = "自覚症状所見"
	cRec[73] = "他覚症状"
	cRec[74] = "他覚症状所見"
	cRec[75] = "保健指導レベル"
	cRec[76] = "服薬・血圧"
	cRec[77] = "服薬・血糖"
	cRec[78] = "服薬・コレステロール"
	cRec[79] = "脳卒中"
	cRec[80] = "心臓病"
	cRec[81] = "慢性腎臓病"
	cRec[82] = "貧血"
	cRec[83] = "たばこ"
	cRec[84] = "体重１０㌔増"
	cRec[85] = "汗かく運動"
	cRec[86] = "歩行１時間以上"
	cRec[87] = "歩く速度"
	cRec[88] = "食事噛む状態"
	cRec[89] = "食べる速度"
	cRec[90] = "就寝前食事"
	cRec[91] = "間食"
	cRec[92] = "朝食抜き"
	cRec[93] = "お酒・頻度"
	cRec[94] = "お酒・量"
	cRec[95] = "睡眠"
	cRec[96] = "改善の意思"
	cRec[97] = "指導受診歴"
	//writer.Write(cRec)
	row := sheet.AddRow()
	for _, cell = range cRec {
		vcell = row.AddCell()
		vcell.Value = cell
	}

	// データ行
	inRecsMax := len(inRecs)
	for J := 1; J < inRecsMax; J++ {
		for I, _ = range cRec {
			cRec[I] = ""
		}

		//　保険証番号が空欄は、データ出力対象外
		if inRecs[J][6] != "" {
			// 0.実施健診機関CD
			cRec[0] = "415201"

			// 1.健診種別CD
			if kazokuCheck(inRecs[J][3]) {
				cRec[1] = "2000" // 家族
			} else {
				cRec[1] = "1000" // 本人
			}

			// 2.受診日
			cRec[2] = strings.Replace((inRecs[J][4]), "-", "/", -1)

			// 3.事業所記号
			cRec[3] = inRecs[J][5]

			// 4.証番号
			cRec[4] = inRecs[J][6]

			// 5.資格区分
			if kazokuCheck(inRecs[J][3]) {
				cRec[5] = "1" // 家族
			} else {
				cRec[5] = "0" // 本人
			}

			// 6.続柄
			cRec[6] = ""

			// 7.枝番
			cRec[7] = ""

			// 8.漢字氏名
			cRec[8] = ""

			// 9.カナ氏名
			cRec[9] = string(norm.NFKC.Bytes([]byte(inRecs[J][7])))

			// 10.性別
			cRec[10] = sei(inRecs[J][8])

			// 11.生年月日
			cRec[11] = WaToSeireki(inRecs[J][9])

			// 12.OP　０１
			cRec[12] = ""

			// 13.OP　０２
			cRec[13] = ""

			// 14.OP　０３
			cRec[14] = ""

			// 15.OP　０４
			cRec[15] = ""

			// 16.OP　０５
			cRec[16] = ""

			// 17.OP　０６
			cRec[17] = ""

			// 18.OP　０７
			cRec[18] = ""

			// 19.OP　０８
			cRec[19] = ""

			// 20.OP　０９
			cRec[20] = ""

			// 21.OP　１０
			cRec[21] = ""

			// 22.OP 11
			cRec[22] = ""

			// 23.請求区分
			cRec[23] = "0"

			// 24.健診金額
			cRec[24] = "7300"

			// 25.法定金額
			cRec[25] = ""

			// 26.請求金額
			cRec[26] = "7300"

			// 27.支払先CD
			cRec[27] = "415201"

			// 28.身長
			cRec[28] = inRecs[J][11]

			// 29.体重
			cRec[29] = inRecs[J][12]

			// 30.BMI
			cRec[30] = inRecs[J][13]

			// 31.腹囲
			cRec[31] = inRecs[J][14]

			// 32.身体検査判定
			cRec[32] = tokkijiko(inRecs[J][44])

			// 33.血圧（収縮期）
			// 34.血圧（拡張期）
			if inRecs[J][15] == "" {
				cRec[33] = ""
				cRec[34] = ""
			} else if inRecs[J][17] == "" {
				cRec[33] = inRecs[J][15]
				cRec[34] = inRecs[J][16]
			} else {
				k1H, _ := strconv.Atoi(inRecs[J][15])
				k1L, _ := strconv.Atoi(inRecs[J][16])
				k2H, _ := strconv.Atoi(inRecs[J][17])
				k2L, _ := strconv.Atoi(inRecs[J][18])
				kH := (k1H + k2H) / 2
				kL := (k1L + k2L) / 2
				cRec[33] = fmt.Sprint(kH)
				cRec[34] = fmt.Sprint(kL)
			}

			// 35.空腹時中性脂肪
			if inRecs[J][179] == "" {
				cRec[35] = inRecs[J][19]
			} else {
				cRec[35] = ""
			}

			// 36.随時中性脂肪
			cRec[36] = inRecs[J][179]

			// 37.HDL・CO
			cRec[37] = inRecs[J][20]

			// 38.LDL・CO
			cRec[38] = inRecs[J][21]

			// 39.Non・HDLCO
			cRec[39] = ""

			// 40.AST(GOT)
			cRec[40] = inRecs[J][22]

			// 41.ALT(GPT)
			cRec[41] = inRecs[J][23]

			// 42.γ・GTP
			cRec[42] = inRecs[J][24]

			// 43.空腹時血糖
			cRec[43] = inRecs[J][25]

			// 44.HｂA1ｃ
			cRec[44] = inRecs[J][26]

			// 45.随時血糖
			cRec[45] = inRecs[J][25]

			// 空腹時血糖・随時血糖の処理
			Eattime, _ := strconv.ParseFloat(inRecs[J][28], 32)
			if (inRecs[J][27] == "とった") && (Eattime < 10) {
				cRec[43] = "" // 随時血糖なので、空腹時血糖の値を空欄にする
			} else {
				cRec[45] = "" // 空腹時血糖なので、随時血糖の値を空欄にする
			}

			// 46.採血時間
			cRec[46] = ""

			// 47.尿糖
			cRec[47] = nyo(inRecs[J][29])

			// 48.尿蛋白
			cRec[48] = nyo(inRecs[J][30])

			// 49.未実施の場合その理由
			cRec[49] = nyoNotReason(inRecs[J][180])

			// 50.白血球数
			cRec[50] = inRecs[J][31]

			// 51.赤血球数
			cRec[51] = inRecs[J][32]

			// 52.血色素量
			cRec[52] = inRecs[J][33]

			// 53.ヘマトクリット
			cRec[53] = inRecs[J][34]

			// 54.心電図所見
			cRec[54] = syokenumu(inRecs[J][76])

			// 55.眼底精密所見
			cRec[55] = syokenumu(inRecs[J][84])

			// 56.血清クレアチニン
			cRec[56] = inRecs[J][35]

			// 57.eGFR
			cRec[57] = inRecs[J][36]

			// 58.HBｓ抗原
			cRec[58] = nyo(inRecs[J][37])

			// 59.HBs抗体
			cRec[59] = nyo(inRecs[J][38])

			// 60.HCV抗体価精密測定
			cRec[60] = nyo(inRecs[J][39])

			// 61.胸部X線検査判定
			cRec[61] = syokenumu(inRecs[J][74])

			// 62.尿酸値
			cRec[62] = inRecs[J][40]

			// 63.腹部超音波検査判定
			cRec[63] = syokenumu(inRecs[J][86])

			// 64.便潜血
			if inRecs[J][42] == "＋" {
				cRec[64] = nyo(inRecs[J][42])
			} else {
				cRec[64] = nyo(inRecs[J][41])
			}

			// 65.総合判定
			//cRec[65] = inRecs[J][43]

			// 66.メタボリック判定
			cRec[66] = ""

			// 67.医師の診断
			// 65.総合判定
			sogo := ""
			var h [7][2]string
			h[0][0] = inRecs[J][44] //身体計測判定
			h[0][1] = inRecs[J][45] //身体計測所見
			h[1][0] = inRecs[J][50] //血圧判定
			h[1][1] = inRecs[J][51] //血圧所見
			if inRecs[J][52] != "" && inRecs[J][53] == "" {
				h[2][0] = inRecs[J][52] //尿蛋白判定
				h[2][1] = inRecs[J][67] //腎機能所見
			} else {
				h[2][0] = inRecs[J][52] //尿蛋白判定
				h[2][1] = inRecs[J][53] //尿蛋白所見
			}
			h[3][0] = inRecs[J][54] //尿糖判定
			h[3][1] = inRecs[J][55] //尿糖所見
			h[4][0] = inRecs[J][60] //血中脂質判定
			h[4][1] = inRecs[J][61] //血中脂質所見
			h[5][0] = inRecs[J][62] //肝機能判定
			h[5][1] = inRecs[J][63] //肝機能所見
			h[6][0] = inRecs[J][64] //糖代謝判定
			h[6][1] = inRecs[J][65] //糖代謝所見

			hKigo := [...]string{"Ｆ", "Ｅ", "Ｄ", "Ｇ", "Ｃ"}
			for k := 0; k < 5; k++ {
				for l := 0; l < 7; l++ {
					if h[l][0] == hKigo[k] {
						if h[l][1] != "" {
							if sogo == "" {
								sogo = h[l][1]
							} else {
								sogo = sogo + "　" + h[l][1]
							}
						}
					}
				}
			}

			cRec[67] = sogo

			sogoHantei := 0
			for l := 0; l < 7; l++ {
				if rank(h[l][0]) > sogoHantei {
					sogoHantei = rank(h[l][0])
				}
			}
			cRec[65] = rankS(sogoHantei)

			// 68.医師名
			cRec[68] = inRecs[J][100]

			// 69.既往歴
			// 70.具体的な既往歴
			kiou := ""
			for k := 0; k < 10; k++ {
				kp := 101 + (k * 3)
				kiouB := kiouSet(inRecs[J][kp])
				kiouN := inRecs[J][kp+1]
				kiouT := inRecs[J][kp+2]

				if kiouB != "" {
					if kiou == "" {
						kiou = kiouB
					} else {
						kiou = kiou + " " + kiouB
					}

					if kiouN != "" {
						kiou = kiou + " " + kiouN + "才"
					}

					if kiouT != "" {
						kiou = kiou + " " + kiouT
					}
				}
			}

			cRec[70] = kiou

			if kiou != "" {
				cRec[69] = "1" // あり
			} else {
				cRec[69] = "2" // なし
			}

			// 71.自覚症状
			// 72.自覚症状所見
			jikaku := ""
			for k := 0; k < 5; k++ {
				kp := 131 + k
				jikakuS := inRecs[J][kp]

				if jikakuS != "" {
					if jikaku == "" {
						jikaku = jikakuS
					} else {
						jikaku = jikaku + " " + jikakuS
					}

				}
			}

			if jikaku == "特になし" {
				jikaku = ""
			}

			cRec[72] = jikaku

			if jikaku != "" {
				cRec[71] = "1" // あり
			} else {
				cRec[71] = "2" // なし
			}

			// 73.他覚症状
			// 74.他覚症状所見
			takaku := ""
			for k := 0; k < 3; k++ {
				kp := 136 + k
				takakuS := inRecs[J][kp]

				if takakuS != "" {
					if takaku == "" {
						takaku = takakuS
					} else {
						takaku = takaku + " " + takakuS
					}

				}
			}

			if takaku == "異常なし" {
				takaku = ""
			}

			cRec[74] = takaku

			if takaku != "" {
				cRec[73] = "1" // あり
			} else {
				cRec[73] = "2" // なし
			}

			// 75.保健指導レベル
			cRec[75] = ""

			// 76.服薬・血圧
			cRec[76] = yesNo(inRecs[J][139])

			// 77.服薬・血糖
			cRec[77] = yesNo(inRecs[J][140])

			// 78.服薬・コレステロール
			cRec[78] = yesNo(inRecs[J][141])

			// 79.脳卒中
			cRec[79] = yesNo(inRecs[J][142])

			// 80.心臓病
			cRec[80] = yesNo(inRecs[J][143])

			// 81.慢性腎臓病
			cRec[81] = yesNo(inRecs[J][144])

			// 82.貧血
			cRec[82] = yesNo(inRecs[J][145])

			// 83.たばこ
			cRec[83] = tabako(inRecs[J][146])

			// 84.体重１０㌔増
			cRec[84] = yesNo(inRecs[J][147])

			// 85.汗かく運動
			cRec[85] = yesNo(inRecs[J][148])

			// 86.歩行１時間以上
			cRec[86] = yesNo(inRecs[J][149])

			// 87.歩く速度
			cRec[87] = yesNo(inRecs[J][150])

			// 88.食事噛む状態
			cRec[88] = eat2(inRecs[J][151])

			// 89.食べる速度
			cRec[89] = eat(inRecs[J][152])

			// 90.就寝前食事
			cRec[90] = yesNo(inRecs[J][153])

			// 91.間食
			cRec[91] = drink(inRecs[J][154])

			// 92.朝食抜き
			cRec[92] = yesNo(inRecs[J][155])

			// 93.お酒・頻度
			cRec[93] = sake(inRecs[J][156])

			// 94.お酒・量
			cRec[94] = sakeryo(inRecs[J][157])

			// 95.睡眠
			cRec[95] = yesNo(inRecs[J][158])

			// 96.改善の意思
			cRec[96] = seikatsu(inRecs[J][159])

			// 97.指導受診歴
			cRec[97] = yesNo(inRecs[J][160])

			//writer.Write(cRec)
			row = sheet.AddRow()
			for _, cell = range cRec {
				// sheet.Cell(r, c).Value = cell
				vcell = row.AddCell()
				vcell.Value = cell
			}
		}
	}

	//writer.Flush()
	err = excelFile.Save(excelName)
	failOnError(err)
}

func gastricConversion(filename string, inRecs [][]string) {
	var vcell *xlsx.Cell
	var cell string

	recLen := 11 //出力する項目数
	cRec := make([]string, recLen)
	var I int

	day := time.Now()

	excelName, _ := filepath.Split(filename)
	excelName = excelName + "松英会職員胃がん検診データ" + day.Format("20060102") + ".xlsx"
	excelFile := xlsx.NewFile()
	xlsx.SetDefaultFont(11, "游ゴシック")
	sheet, err := excelFile.AddSheet("データ")
	failOnError(err)

	//タイトル行
	cRec[0] = "支払先CD"
	cRec[1] = "受診日"
	cRec[2] = "事業所記号"
	cRec[3] = "証番号"
	cRec[4] = "資格区分"
	cRec[5] = "カナ氏名"
	cRec[6] = "性別"
	cRec[7] = "生年月日"
	cRec[8] = "結果"
	cRec[9] = "所見"
	cRec[10] = "検査区分"
	//writer.Write(cRec)
	row := sheet.AddRow()
	for _, cell = range cRec {
		vcell = row.AddCell()
		vcell.Value = cell
	}

	// データ行
	inRecsMax := len(inRecs)
	for J := 1; J < inRecsMax; J++ {
		for I, _ = range cRec {
			cRec[I] = ""
		}

		//　保険証番号が空欄は、データ出力対象外
		if inRecs[J][6] != "" {
			if inRecs[J][78] != "" || inRecs[J][80] != "" {
				// 0.支払先CD
				cRec[0] = "415201"

				// 1.受診日
				cRec[1] = strings.Replace((inRecs[J][4]), "-", "/", -1)

				// 2.事業所記号
				cRec[2] = inRecs[J][5]

				// 3.証番号
				cRec[3] = inRecs[J][6]

				// 4.資格区分
				if kazokuCheck(inRecs[J][3]) {
					cRec[4] = "1" // 家族
				} else {
					cRec[4] = "0" // 本人
				}

				// 5.カナ氏名
				cRec[5] = string(norm.NFKC.Bytes([]byte(inRecs[J][7])))

				// 6.性別
				cRec[6] = sei(inRecs[J][8])

				// 7.生年月日
				cRec[7] = WaToSeireki(inRecs[J][9])

				// 8.結果区分
				cRec[8] = kekka(inRecs[J][78])

				// 9.所見

				//胃部X線か胃カメラか
				syoken := ""
				if inRecs[J][78] != "" && inRecs[J][78] != "Ａ" && inRecs[J][78] != "Ｂ" {
					// 胃部X線の場合
					for k := 0; k < 3; k++ {
						kp := 161 + k
						syokenS := inRecs[J][kp]

						if syokenS != "" {
							if syoken == "" {
								syoken = syokenS
							} else {
								syoken = syoken + " " + syokenS
							}

						}
					}
				} else if inRecs[J][80] != "" {
					// 胃カメラの場合
					syoken := ""
					for k := 0; k < 3; k++ {
						kp := 164 + k
						syokenS := inRecs[J][kp]

						if syokenS != "" {
							if syoken == "" {
								syoken = syokenS
							} else {
								syoken = syoken + " " + syokenS
							}

						}
					}
				}

				cRec[9] = syoken

				// 10.結果
				cRec[10] = "レントゲン"

				//writer.Write(cRec)
				row = sheet.AddRow()
				for _, cell = range cRec {
					// sheet.Cell(r, c).Value = cell
					vcell = row.AddCell()
					vcell.Value = cell
				}
			}
		}
	}

	//writer.Flush()
	err = excelFile.Save(excelName)
	failOnError(err)
}

func uterineConversion(filename string, inRecs [][]string) {
	var vcell *xlsx.Cell
	var cell string

	recLen := 11 //出力する項目数
	cRec := make([]string, recLen)
	var I int

	day := time.Now()

	excelName, _ := filepath.Split(filename)
	excelName = excelName + "松英会職員子宮がん検診データ" + day.Format("20060102") + ".xlsx"
	excelFile := xlsx.NewFile()
	xlsx.SetDefaultFont(11, "游ゴシック")
	sheet, err := excelFile.AddSheet("データ")
	failOnError(err)

	//タイトル行
	cRec[0] = "支払先CD"
	cRec[1] = "受診日"
	cRec[2] = "事業所記号"
	cRec[3] = "証番号"
	cRec[4] = "資格区分"
	cRec[5] = "カナ氏名"
	cRec[6] = "性別"
	cRec[7] = "生年月日"
	cRec[8] = "結果"
	cRec[9] = "所見"
	cRec[10] = "検査区分"
	//writer.Write(cRec)
	row := sheet.AddRow()
	for _, cell = range cRec {
		vcell = row.AddCell()
		vcell.Value = cell
	}

	// データ行
	inRecsMax := len(inRecs)
	for J := 1; J < inRecsMax; J++ {
		for I, _ = range cRec {
			cRec[I] = ""
		}

		//　保険証番号が空欄は、データ出力対象外
		if inRecs[J][6] != "" {
			if inRecs[J][90] != "" {
				// 0.支払先CD
				cRec[0] = "415201"

				// 1.受診日
				cRec[1] = strings.Replace((inRecs[J][4]), "-", "/", -1)

				// 2.事業所記号
				cRec[2] = inRecs[J][5]

				// 3.証番号
				cRec[3] = inRecs[J][6]

				// 4.資格区分
				if kazokuCheck(inRecs[J][3]) {
					cRec[4] = "1" // 家族
				} else {
					cRec[4] = "0" // 本人
				}

				// 5.カナ氏名
				cRec[5] = string(norm.NFKC.Bytes([]byte(inRecs[J][7])))

				// 6.性別
				cRec[6] = sei(inRecs[J][8])

				// 7.生年月日
				cRec[7] = WaToSeireki(inRecs[J][9])

				// 8.結果
				cRec[8] = kekka(inRecs[J][90])

				// 9.所見

				cRec[9] = ""

				// 10.検査区分
				cRec[10] = ""

				//writer.Write(cRec)
				row = sheet.AddRow()
				for _, cell = range cRec {
					// sheet.Cell(r, c).Value = cell
					vcell = row.AddCell()
					vcell.Value = cell
				}
			}
		}
	}

	//writer.Flush()
	err = excelFile.Save(excelName)
	failOnError(err)
}

func breastConversion(filename string, inRecs [][]string) {
	var vcell *xlsx.Cell
	var cell string

	recLen := 11 //出力する項目数
	cRec := make([]string, recLen)
	var I int

	day := time.Now()

	excelName, _ := filepath.Split(filename)
	excelName = excelName + "松英会職員乳がん検診データ" + day.Format("20060102") + ".xlsx"
	excelFile := xlsx.NewFile()
	xlsx.SetDefaultFont(11, "游ゴシック")
	sheet, err := excelFile.AddSheet("データ")
	failOnError(err)

	//タイトル行
	cRec[0] = "支払先CD"
	cRec[1] = "受診日"
	cRec[2] = "事業所記号"
	cRec[3] = "証番号"
	cRec[4] = "資格区分"
	cRec[5] = "カナ氏名"
	cRec[6] = "性別"
	cRec[7] = "生年月日"
	cRec[8] = "結果"
	cRec[9] = "所見"
	cRec[10] = "検査区分"
	//writer.Write(cRec)
	row := sheet.AddRow()
	for _, cell = range cRec {
		vcell = row.AddCell()
		vcell.Value = cell
	}

	// データ行
	inRecsMax := len(inRecs)
	for J := 1; J < inRecsMax; J++ {
		for I, _ = range cRec {
			cRec[I] = ""
		}

		//　保険証番号が空欄は、データ出力対象外
		if inRecs[J][6] != "" {
			if inRecs[J][94] != "" {
				// 0.支払先CD
				cRec[0] = "415201"

				// 1.受診日
				cRec[1] = strings.Replace((inRecs[J][4]), "-", "/", -1)

				// 2.事業所記号
				cRec[2] = inRecs[J][5]

				// 3.証番号
				cRec[3] = inRecs[J][6]

				// 4.資格区分
				if kazokuCheck(inRecs[J][3]) {
					cRec[4] = "1" // 家族
				} else {
					cRec[4] = "0" // 本人
				}

				// 5.カナ氏名
				cRec[5] = string(norm.NFKC.Bytes([]byte(inRecs[J][7])))

				// 6.性別
				cRec[6] = sei(inRecs[J][8])

				// 7.生年月日
				cRec[7] = WaToSeireki(inRecs[J][9])

				// 8.結果
				cRec[8] = kekka(inRecs[J][94])

				// 9.所見

				syoken := ""
				if inRecs[J][94] != "" && inRecs[J][94] != "Ａ" && inRecs[J][94] != "Ｂ" {
					for k := 0; k < 3; k++ {
						kp := 171 + k
						syokenS := inRecs[J][kp]

						if syokenS != "" {
							if syoken == "" {
								syoken = syokenS
							} else {
								syoken = syoken + " " + syokenS
							}

						}
					}
				}

				cRec[9] = syoken

				// 10.検査区分
				cRec[10] = "超音波"

				//writer.Write(cRec)
				row = sheet.AddRow()
				for _, cell = range cRec {
					// sheet.Cell(r, c).Value = cell
					vcell = row.AddCell()
					vcell.Value = cell
				}
			}
		}
	}

	//writer.Flush()
	err = excelFile.Save(excelName)
	failOnError(err)
}

func prostateConversion(filename string, inRecs [][]string) {
	var vcell *xlsx.Cell
	var cell string

	recLen := 11 //出力する項目数
	cRec := make([]string, recLen)
	var I int

	day := time.Now()

	excelName, _ := filepath.Split(filename)
	excelName = excelName + "松英会職員前立腺がん検診データ" + day.Format("20060102") + ".xlsx"
	excelFile := xlsx.NewFile()
	xlsx.SetDefaultFont(11, "游ゴシック")
	sheet, err := excelFile.AddSheet("データ")
	failOnError(err)

	//タイトル行
	cRec[0] = "支払先CD"
	cRec[1] = "受診日"
	cRec[2] = "事業所記号"
	cRec[3] = "証番号"
	cRec[4] = "資格区分"
	cRec[5] = "カナ氏名"
	cRec[6] = "性別"
	cRec[7] = "生年月日"
	cRec[8] = "結果"
	cRec[9] = "所見"
	cRec[10] = "検査区分"
	//writer.Write(cRec)
	row := sheet.AddRow()
	for _, cell = range cRec {
		vcell = row.AddCell()
		vcell.Value = cell
	}

	// データ行
	inRecsMax := len(inRecs)
	for J := 1; J < inRecsMax; J++ {
		for I, _ = range cRec {
			cRec[I] = ""
		}

		//　保険証番号が空欄は、データ出力対象外
		if inRecs[J][6] != "" {
			if inRecs[J][175] != "" {
				// 0.支払先CD
				cRec[0] = "415201"

				// 1.受診日
				cRec[1] = strings.Replace((inRecs[J][4]), "-", "/", -1)

				// 2.事業所記号
				cRec[2] = inRecs[J][5]

				// 3.証番号
				cRec[3] = inRecs[J][6]

				// 4.資格区分
				if kazokuCheck(inRecs[J][3]) {
					cRec[4] = "1" // 家族
				} else {
					cRec[4] = "0" // 本人
				}

				// 5.カナ氏名
				cRec[5] = string(norm.NFKC.Bytes([]byte(inRecs[J][7])))

				// 6.性別
				cRec[6] = sei(inRecs[J][8])

				// 7.生年月日
				cRec[7] = WaToSeireki(inRecs[J][9])

				// 8.結果
				cRec[8] = kekka(inRecs[J][175])

				// 9.所見

				cRec[9] = "PSA " + inRecs[J][174]

				// 10.検査区分
				cRec[10] = ""

				//writer.Write(cRec)
				row = sheet.AddRow()
				for _, cell = range cRec {
					// sheet.Cell(r, c).Value = cell
					vcell = row.AddCell()
					vcell.Value = cell
				}
			}
		}
	}

	//writer.Flush()
	err = excelFile.Save(excelName)
	failOnError(err)
}

func mmgConversion(filename string, inRecs [][]string) {
	var vcell *xlsx.Cell
	var cell string

	recLen := 11 //出力する項目数
	cRec := make([]string, recLen)
	var I int

	day := time.Now()

	excelName, _ := filepath.Split(filename)
	excelName = excelName + "松英会職員マンモ検診データ" + day.Format("20060102") + ".xlsx"
	excelFile := xlsx.NewFile()
	xlsx.SetDefaultFont(11, "游ゴシック")
	sheet, err := excelFile.AddSheet("データ")
	failOnError(err)

	//タイトル行
	cRec[0] = "支払先CD"
	cRec[1] = "受診日"
	cRec[2] = "事業所記号"
	cRec[3] = "証番号"
	cRec[4] = "資格区分"
	cRec[5] = "カナ氏名"
	cRec[6] = "性別"
	cRec[7] = "生年月日"
	cRec[8] = "結果"
	cRec[9] = "所見"
	cRec[10] = "検査区分"
	//writer.Write(cRec)
	row := sheet.AddRow()
	for _, cell = range cRec {
		vcell = row.AddCell()
		vcell.Value = cell
	}

	// データ行
	inRecsMax := len(inRecs)
	for J := 1; J < inRecsMax; J++ {
		for I, _ = range cRec {
			cRec[I] = ""
		}

		//　保険証番号が空欄は、データ出力対象外
		if inRecs[J][6] != "" {
			if inRecs[J][96] != "" {
				// 0.支払先CD
				cRec[0] = "415201"

				// 1.受診日
				cRec[1] = strings.Replace((inRecs[J][4]), "-", "/", -1)

				// 2.事業所記号
				cRec[2] = inRecs[J][5]

				// 3.証番号
				cRec[3] = inRecs[J][6]

				// 4.資格区分
				if kazokuCheck(inRecs[J][3]) {
					cRec[4] = "1" // 家族
				} else {
					cRec[4] = "0" // 本人
				}

				// 5.カナ氏名
				cRec[5] = string(norm.NFKC.Bytes([]byte(inRecs[J][7])))

				// 6.性別
				cRec[6] = sei(inRecs[J][8])

				// 7.生年月日
				cRec[7] = WaToSeireki(inRecs[J][9])

				// 8.結果
				cRec[8] = kekka(inRecs[J][96])

				// 9.所見

				syoken := ""
				if inRecs[J][96] != "" && inRecs[J][96] != "Ａ" && inRecs[J][96] != "Ｂ" {
					for k := 0; k < 3; k++ {
						kp := 176 + k
						syokenS := inRecs[J][kp]

						if syokenS != "" {
							if syoken == "" {
								syoken = syokenS
							} else {
								syoken = syoken + " " + syokenS
							}

						}
					}
				}

				cRec[9] = syoken

				// 10.検査区分
				cRec[10] = "マンモ"

				//writer.Write(cRec)
				row = sheet.AddRow()
				for _, cell = range cRec {
					// sheet.Cell(r, c).Value = cell
					vcell = row.AddCell()
					vcell.Value = cell
				}
			}
		}
	}

	//writer.Flush()
	err = excelFile.Save(excelName)
	failOnError(err)
}

func dexaConversion(filename string, inRecs [][]string) {
	var vcell *xlsx.Cell
	var cell string

	recLen := 7 //出力する項目数
	cRec := make([]string, recLen)
	var I int

	day := time.Now()

	excelName, _ := filepath.Split(filename)
	excelName = excelName + "松英会職員骨密度検診データ" + day.Format("20060102") + ".xlsx"
	excelFile := xlsx.NewFile()
	xlsx.SetDefaultFont(11, "游ゴシック")
	sheet, err := excelFile.AddSheet("データ")
	failOnError(err)

	//タイトル行
	cRec[0] = "利用日"
	cRec[1] = "記号"
	cRec[2] = "番号"
	cRec[3] = "本人家族"
	cRec[4] = "カナ氏名"
	cRec[5] = "生年月日"
	cRec[6] = "実施金額"
	//writer.Write(cRec)
	row := sheet.AddRow()
	for _, cell = range cRec {
		vcell = row.AddCell()
		vcell.Value = cell
	}

	// データ行
	inRecsMax := len(inRecs)
	for J := 1; J < inRecsMax; J++ {
		for I, _ = range cRec {
			cRec[I] = ""
		}

		//　保険証番号が空欄は、データ出力対象外
		if inRecs[J][6] != "" {
			if inRecs[J][181] != "" {
				// 0.利用日
				cRec[0] = strings.Replace((inRecs[J][4]), "-", "/", -1)

				// 1.記号
				cRec[1] = inRecs[J][5]

				// 2.番号
				cRec[2] = inRecs[J][6]

				// 3.本人家族
				if kazokuCheck(inRecs[J][3]) {
					cRec[3] = "1" // 家族
				} else {
					cRec[3] = "0" // 本人
				}

				// 4.カナ氏名（半角）
				cRec[4] = inRecs[J][7]

				// 5.生年月日
				cRec[5] = WaToSeireki(inRecs[J][9])

				// 6.実施金額
				cRec[6] = "3000"

				//writer.Write(cRec)
				row = sheet.AddRow()
				for _, cell = range cRec {
					// sheet.Cell(r, c).Value = cell
					vcell = row.AddCell()
					vcell.Value = cell
				}
			}
		}
	}

	//writer.Flush()
	err = excelFile.Save(excelName)
	failOnError(err)
}

func WaToSeireki(nen string) string {

	if len(nen) != 9 {
		return nen
	} else {
		w := nen[0:1]
		y := nen[1 : 1+2]
		yi, _ := strconv.Atoi(y)
		m := nen[4 : 4+2]
		d := nen[7 : 7+2]

		switch w {
		case "M":
			yi = yi + 1867
		case "T":
			yi = yi + 1911
		case "S":
			yi = yi + 1925
		case "H":
			yi = yi + 1988
		default:
			yi = 0
		}

		if yi == 0 {
			return "err"
		} else {
			return fmt.Sprint(yi) + "/" + m + "/" + d
		}
	}
}

func sei(s string) string {

	if s == "男" {
		return "1"
	} else if s == "女" {
		return "2"
	} else {
		log.Print("性別エラー\r\n")
		return "err"
	}
}

func kiouSet(s string) string {
	var spos, epos int
	//全角記号を半角へ
	s = strings.Replace(s, "（", "(", -1)
	s = strings.Replace(s, "）", ")", -1)
	s = strings.Replace(s, "　", " ", -1)

	// ()でくくった文字は削除
	for {
		spos = strings.LastIndex(s, "(")
		epos = strings.LastIndex(s, ")")

		if epos == -1 {
			break
		} else if spos == -1 {
			break
		} else {
			//log.Print(s + ":epos→" + fmt.Sprint(epos) + " len→" + fmt.Sprint(len(s)) + "\r\n")
			s = s[:spos] + s[epos+1:]
		}
	}

	// 余分なスペースを削除
	s = dsTrim(s)
	s = strings.Trim(s, " ")

	return s
}

func dsTrim(s string) string {
	for {
		if strings.Contains(s, "  ") {
			s = strings.Replace(s, "  ", " ", -1)
		} else {
			return s
		}
	}
}

func cutStrings(s string, maxLen int) string {
	s = string([]rune(s)[:maxLen])
	return s
}

func syoken(s string) string {
	s = strings.Replace(s, "　", " ", -1)
	s = strings.Trim(s, " ")

	for {
		if utf8.RuneCountInString(s) > 25 {
			pos := strings.LastIndex(s, " ")
			s = s[:pos]
		} else {
			break
		}
	}

	return s
}

func nyo(s string) string {

	switch s {
	case "":
		s = ""
	case "－":
		s = "-"
	case "+-":
		s = "+-"
	case "＋":
		s = "+"
	case "2+":
		s = "++"
	case "3+":
		s = "+++"
	case "4+":
		s = "+++"
	case "5+":
		s = "+++"
	default:
		log.Print("尿変換エラー\r\n")
		s = "err"
	}

	return s
}

func tokkijiko(s string) string {
	// 特記事項あり:1
	// 特記事項なし:2

	switch s {
	case "":
		s = ""
	case "Ａ":
		s = "2"
	case "Ｂ":
		s = "2"
	case "Ｃ":
		s = "1"
	case "Ｄ":
		s = "1"
	case "Ｅ":
		s = "1"
	case "Ｆ":
		s = "1"
	case "Ｇ":
		s = "1"
	default:
		log.Print("判定有無変換エラー\r\n")
		s = "err"
	}
	return s
}

func syokenumu(s string) string {
	// 所見あり:1
	// 所見なし:2

	switch s {
	case "":
		s = ""
	case "Ａ":
		s = "2"
	case "Ｂ":
		s = "1"
	case "Ｃ":
		s = "1"
	case "Ｄ":
		s = "1"
	case "Ｅ":
		s = "1"
	case "Ｆ":
		s = "1"
	case "Ｇ":
		s = "1"
	case "Ｈ":
		s = "1"
	default:
		log.Print("判定有無変換エラー\r\n")
		s = "err"
	}
	return s
}

func kekka(s string) string {
	switch s {
	case "":
		s = ""
	case "Ａ":
		s = "1"
	case "Ｂ":
		s = "2"
	case "Ｃ":
		s = "3"
	case "Ｄ":
		s = "4"
	case "Ｅ":
		s = "4"
	case "Ｆ":
		s = "5"
	case "Ｇ":
		s = "6"
	case "A":
		s = "1"
	case "B":
		s = "2"
	case "C":
		s = "3"
	case "D":
		s = "4"
	case "E":
		s = "4"
	case "F":
		s = "5"
	case "G":
		s = "6"
	default:
		log.Print("結果区分変換エラー\r\n")
		s = "err"
	}
	return s
}

func rank(v string) int {
	r := 0
	switch v {
	case "":
		r = 0
	case "Ａ":
		r = 1
	case "Ｂ":
		r = 2
	case "Ｃ":
		r = 3
	case "Ｄ":
		r = 5
	case "Ｅ":
		r = 6
	case "Ｆ":
		r = 7
	case "Ｇ":
		r = 4
	default:
		r = 0
		//log.Print("判定ランクエラー\r\n")
		log.Printf("判定ランクエラーa%sa\r\n",v)
	}
	return r
}

func rankS(v int) string {
	r := ""
	switch v {
	case 1:
		r = "所見なし"
	case 2:
		r = "略正常"
	case 3:
		r = "要観察"
	case 4:
		r = "治療中"
	case 5:
		r = "要再検"
	case 6:
		r = "要再検"
	case 7:
		r = "要治療"
	default:
		r = "err"
		log.Print("判定ランクコメントエラー\r\n")
	}
	return r
}

func kazokuCheck(v string) bool {
	// 職員家族なら true

	if v == "職員家族" {
		return true
	} else {
		return false
	}
}

func yesNo(s string) string {

	switch s {
	case "":
		s = ""
	case "はい":
		s = "1"
	case "いいえ":
		s = "2"
	default:
		log.Print("はいいいえ変換エラー\r\n")
		s = "err"
	}
	return s
}

func eat(s string) string {

	switch s {
	case "":
		s = ""
	case "速い":
		s = "1"
	case "普通":
		s = "2"
	case "遅い":
		s = "3"
	default:
		s = "err"
	}
	return s
}

func eat2(s string) string {

	switch s {
	case "":
		s = ""
	case "何でも":
		s = "1"
	case "かみにくい":
		s = "2"
	case "ほとんどかめない":
		s = "3"
	default:
		log.Print("かんで食べる変換エラー\r\n")
		s = "err"
	}
	return s
}

func drink(s string) string {

	switch s {
	case "":
		s = ""
	case "毎日":
		s = "1"
	case "時々":
		s = "2"
	case "ほとんど摂取しない":
		s = "3"
	default:
		log.Print("間食あまい飲み物変換エラー\r\n")
		s = "err"
	}
	return s
}

func sake(s string) string {

	switch s {
	case "":
		s = ""
	case "毎日":
		s = "1"
	case "週５～６日":
		s = "2"
	case "週３～４日":
		s = "3"
	case "週１～２日":
		s = "4"
	case "月に１～３日":
		s = "5"
	case "月に１日未満":
		s = "6"
	case "やめた":
		s = "7"
	case "飲まない":
		s = "8"
	default:
		log.Print("お酒変換エラー\r\n")
		s = "err"
	}
	return s
}

func sakeryo(s string) string {

	switch s {
	case "":
		s = ""
	case "１合未満":
		s = "1"
	case "１～２合未満":
		s = "2"
	case "２～３合未満":
		s = "3"
	case "３～５合未満":
		s = "4"
	case "５合以上":
		s = "5"
	default:
		log.Print("飲酒量変換エラー\r\n")
		s = "err"
	}
	return s
}

func seikatsu(s string) string {

	switch s {
	case "":
		s = ""
	case "しない":
		s = "1"
	case "思う":
		s = "2"
	case "始めた":
		s = "3"
	case "６ヶ月経過":
		s = "4"
	case "６ヶ月以上":
		s = "5"
	default:
		log.Print("生活習慣改善変換エラー\r\n")
		s = "err"
	}
	return s
}

func nyoNotReason(s string) string {

	switch s {
	case "":
		s = ""
	case "生理中":
		s = "1"
	case "腎疾患等の基礎疾患があるため排尿障害を有する":
		s = "2"
	case "その他":
		s = "3"
	default:
		log.Print("未実施の場合その理由変換エラー\r\n")
		s = "err"
	}
	return s
}

func tabako(s string) string {
	
	switch s {
	case "":
		s = ""
	case "はい":
		s = "1"
	case "以前あり":
		s ="2"
	case "いいえ":
		s = "3"
	default:
		log.Print("たばこ変換エラー\r\n")
		s = "err"
	}
	return s
}

