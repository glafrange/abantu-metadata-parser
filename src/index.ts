import { ColInfo } from "xlsx"
import * as xlsx from "xlsx"
import * as fs from "fs/promises"
import * as fsSync from "fs"

import { AdditionalBookInfo, BisacRow, BookMetadata, MetadataParser } from "./metadata-parser"
import { entriesFromObject } from "./utils"

const OUTPUT_SHEET_PATH = "./output/book-metadata.xlsx"
const OUTPUT_XML_PATH = "./output_xml/"

async function main() {
  const option = process.argv[2]
  if (option === "parse") {
    await parseMetadata()
  }
  else if (option === "collect") {
    await collectMetadataFiles()
  }
}

async function collectMetadataFiles() {
  const metadataWorkbook = xlsx.readFile('./output/book-metadata.xlsx')
  const metadata: BookMetadata[] = xlsx.utils.sheet_to_json(metadataWorkbook.Sheets['Sheet1'])
  const books = metadata.filter(book => book["To Collect (x)"]?.length > 0)
  const bookInfoFile = await fs.readFile("./data/additional-book-info.json", 'utf8')
  const bookInfoMap: Map<string, AdditionalBookInfo> = new Map(JSON.parse(bookInfoFile))
  const bookAndInfo = books.reduce((acc, book) => {
    const bookInfo = bookInfoMap.get(book.isbn)
    if (!book) return acc
    if (!bookInfo) return acc
    return [...acc, [bookInfo, book]] as [AdditionalBookInfo, BookMetadata][]
  }, [] as [AdditionalBookInfo, BookMetadata][])
  bookAndInfo.forEach((arr) => {
    const [bookInfo, book] = arr
    fsSync.copyFileSync(`${bookInfo.originFilePath}`, `${OUTPUT_XML_PATH}${book.isbn}.xml`)
    bookInfo.toCollect = ""
    fsSync.writeFileSync(`${OUTPUT_XML_PATH}${book.isbn}.json`, JSON.stringify(book))
  })  
  const newBooks = metadata.map(book => {
    book["To Collect (x)"] = bookInfoMap.get(book.isbn)?.toCollect ?? ""
    return book
  })
  const columnNames = Object.keys(newBooks[0])
  let newSheet = xlsx.utils.json_to_sheet(newBooks)
  newSheet = xlsx.utils.sheet_add_aoa(newSheet, [[...columnNames]], { origin: "A1" })
  addCellOptions(newSheet, columnNames, newBooks)
  let workBook = xlsx.utils.book_new()
  xlsx.utils.book_append_sheet(workBook, newSheet)
  xlsx.writeFile(workBook, OUTPUT_SHEET_PATH, { cellStyles: true })

  console.log(`${books.length} books collected in ${OUTPUT_XML_PATH}`)
}

async function parseMetadata() {
  const bisacWorkbook = xlsx.readFile('bisac.xlsx')
  const bisacCodeSheet = bisacWorkbook.Sheets['bisac code']
  const subjectHeadingSheet = bisacWorkbook.Sheets['Subject Heading Text']
  const bisacSheet: BisacRow[] = xlsx.utils.sheet_to_json(bisacCodeSheet)
  const mParser = new MetadataParser(subjectHeadingSheet, bisacSheet)
  const bookList = await mParser.parse()
  const columnNames = Object.keys(bookList[0])
  let metadataSheet = xlsx.utils.json_to_sheet(bookList)
  metadataSheet = xlsx.utils.sheet_add_aoa(metadataSheet, [[...columnNames]], { origin: "A1" })
  addCellOptions(metadataSheet, columnNames, bookList)
  let workBook = xlsx.utils.book_new()
  xlsx.utils.book_append_sheet(workBook, metadataSheet)
  xlsx.writeFile(workBook, OUTPUT_SHEET_PATH, { cellStyles: true })
  console.log(`File Created: ${OUTPUT_SHEET_PATH}`)
}

function addCellOptions(sheet: xlsx.WorkSheet, columnNames: string[], bookList: BookMetadata[]) {
  const defaultInfo = {
    wpx: 150
  }
  const columnInfo: Record<keyof BookMetadata, ColInfo> = {
    isbn: {
      wpx: defaultInfo.wpx
    },
    title: {
      wpx: defaultInfo.wpx
    },
    subtitle: {
      wpx: defaultInfo.wpx
    },
    contributors: {
      wpx: defaultInfo.wpx
    },
    imprint: {
      wpx: defaultInfo.wpx
    },
    pubDate: {
      wpx: defaultInfo.wpx
    },
    onSaleDate: {
      wpx: defaultInfo.wpx
    },
    usPrice: {
      wpx: defaultInfo.wpx
    },
    caPrice: {
      wpx: defaultInfo.wpx
    },
    runtime: {
      wpx: defaultInfo.wpx
    },
    BISAC: {
      wpx: defaultInfo.wpx
    },
    language: {
      wpx: defaultInfo.wpx
    },
    primaryCategory: {
      wpx: defaultInfo.wpx
    },
    secondaryCategories: {
      wpx: defaultInfo.wpx
    },
    customCategory: {
      wpx: defaultInfo.wpx
    },
    "To Collect (x)": {
      wpx: defaultInfo.wpx
    }
  } as const
  const entries = entriesFromObject(columnInfo)
  for (let i = 0; i < entries.length; ++i) {
    const [colName, colInfo] = entries[i]
    if (!sheet["!cols"]) sheet["!cols"] = []
    if (!sheet["!cols"][i]) sheet["!cols"][i] = columnInfo[colName]
  }
}

main()