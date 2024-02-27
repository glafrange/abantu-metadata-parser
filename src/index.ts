import { ColInfo } from "xlsx"
import * as xlsx from "xlsx"
import * as fs from "fs/promises"

import { BisacRow, BookMetadata, MetadataParser } from "./metadata-parser"

const PUBLISHER_DIRECTORIES = ["SimonSchuster", "Hachette"] as const // directories containing the xml files
const OUTPUT_PATH = "./output/book-metadata.xlsx"

type Entries<T> = {
  [K in keyof T]: [K, T[K]];
}[keyof T][];

function entriesFromObject<T extends object>(object: T): Entries<T> {
  return Object.entries(object) as Entries<T>;
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
    }
  } as const
  const entries = entriesFromObject(columnInfo)
 for (let i = 0; i < entries.length; ++i) {
    const [colName, colInfo] = entries[i]
    if (!sheet["!cols"]) sheet["!cols"] = []
    if (!sheet["!cols"][i]) sheet["!cols"][i] = columnInfo[colName]
  }
}


async function main() {
  const bisacWorkbook = xlsx.readFile('bisac.xlsx')
  const bisacCodeSheet = bisacWorkbook.Sheets['bisac code']
  const subjectHeadingSheet = bisacWorkbook.Sheets['Subject Heading Text']
  const bisacSheet: BisacRow[] = xlsx.utils.sheet_to_json(bisacCodeSheet)
  const mParser = new MetadataParser(PUBLISHER_DIRECTORIES, subjectHeadingSheet, bisacSheet)
  const bookList = await mParser.parse()
  const columnNames = Object.keys(bookList[0])
  let metadataSheet = xlsx.utils.json_to_sheet(bookList)
  metadataSheet = xlsx.utils.sheet_add_aoa(metadataSheet, [[...columnNames]], { origin: "A1" })
  addCellOptions(metadataSheet, columnNames, bookList)
  let workBook = xlsx.utils.book_new()
  xlsx.utils.book_append_sheet(workBook, metadataSheet)
  xlsx.writeFile(workBook, OUTPUT_PATH, { cellStyles: true })
  console.log(`File Created: ${OUTPUT_PATH}`)
}

main()