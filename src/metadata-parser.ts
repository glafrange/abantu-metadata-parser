import * as fs from 'fs/promises'
import { Element } from "elementtree"
import * as ET from "elementtree"
import { z } from "zod"
import { WorkSheet } from 'xlsx'
import * as xlsx from 'xlsx'

// data related to a books parsing
type AdditionalBookInfo = {
  originFile: string,
  onixVer: number
}
export const additionalBookInfo = new WeakMap<Object, AdditionalBookInfo>()

export type BisacRow = {
  "BISAC  Code": string,
  "Category": string,
  "Subject Flag"?: string,
  "Custom Category"?: string
}

export const BookMetadataSchema = z.object({
  isbn: z.string(),
  title: z.string(),
  subtitle: z.optional(z.string()),
  contributors: z.string(),
  imprint: z.string(),
  pubDate: z.string(),
  onSaleDate: z.string(),
  usPrice: z.string(),
  caPrice: z.optional(z.string()),
  runtime: z.optional(z.string()),
  BISAC: z.string(),
  language: z.string(),
  primaryCategory: z.optional(z.string()),
  secondaryCategories: z.optional(z.string()),
  customCategory: z.optional(z.string()),
})
export type BookMetadata = z.infer<typeof BookMetadataSchema>

export class MetadataParser {
  private subjectPhrases: string[]

  constructor(
    private pubDirs: Readonly<string[]>, 
    subjectHeadingSheet: WorkSheet, 
    private bisacSheet: BisacRow[],
  ) {
    this.subjectPhrases = xlsx.utils.sheet_to_json(subjectHeadingSheet)
    this.subjectPhrases.splice(0, 1)
  }

  private async getRawMetadataList() {
    type RawMetadata = {
      onixVer: number,
      productElement: Element,
      originFile: string,
    }
    const xmlDirs = await fs.readdir(`xml_metadata`)
    const metadataList = await xmlDirs.reduce(async(accPromise, dir) => {
      const acc = await accPromise
      const pathList = (await fs.readdir(`xml_metadata/${dir}`)).filter(fileName => fileName.endsWith(".xml"))     
      const metadataList = await pathList.reduce(async(accPromise, fileName) => {
        const acc = await accPromise
        const xmlString = await fs.readFile(`xml_metadata/${dir}/${fileName}`, 'utf-8')
        const root = ET.parse(xmlString).getroot()
        const product = root.findall('Product')

        if(product.length > 1) {
          const rawMetadataList: RawMetadata[] = product.map((product) => {
            return {
              onixVer: parseInt(root.get('release') ?? '0'), 
              productElement: product,
              originFile: fileName
            }
          })
          return [...acc, ...rawMetadataList]
        } else {
          const rawMetadata: RawMetadata = {
            onixVer: parseInt(root.get('release') ?? '2.1'), //single product files assumed Hachette & ONIX 2.1
            productElement: product[0],
            originFile: fileName
          }
          return [...acc, rawMetadata]
        }
      }, Promise.resolve([] as RawMetadata[]))
      return [...acc, ...metadataList]
    }, Promise.resolve([] as RawMetadata[]))
    return metadataList
  }

  private formatRuntime(extentUnit: string | undefined, runtime: string | undefined): string | undefined {
    if(!extentUnit || !runtime) return undefined
    if (extentUnit == "16") {
     return `${runtime.slice(0, 3)}:${runtime.slice(3,5)}:${runtime.slice(5)}`
    } else if (extentUnit === "05") {
      const floatRuntime = parseFloat(runtime)
      const hours = Math.floor(floatRuntime / 60)
      const minutes = Math.floor(floatRuntime % 60)
      return `${hours}:${minutes}`
    } else {
      return undefined
    }
  }

  private matchSubjectText(subjectText: string): string {
    let customCategoryList: string[] = []
    this.subjectPhrases.forEach(phrase => {
      const [text, category] = Object.values(phrase)
      if (subjectText.includes(text) && !customCategoryList.includes(category)) {
        customCategoryList.push(category)
      }
    })
    return customCategoryList.join(", ")
  }

  private getCategories(bisac: string): {
    primaryCategory?: string, 
    secondaryCategory?: string
  } {
    const bisacRow: BisacRow | undefined = this.bisacSheet.find((row: BisacRow) => {
      return row["BISAC  Code"] === bisac
    })
    if (!bisacRow) return { primaryCategory: "", secondaryCategory: "" }
    const categories = bisacRow.Category.split(" / ")
    return {
      primaryCategory: categories.splice(0, 1)[0],
      secondaryCategory: categories.join(" / "),
    }
  }

  public async parse(): Promise<BookMetadata[]> {
    const rawMetadataList = await this.getRawMetadataList()
    if (!rawMetadataList) return []
    const parsedMetadata = rawMetadataList.reduce(async (accPromise, { onixVer, productElement, originFile }) => {
      const acc = await accPromise
      const extentElem = productElement.findall(".//Extent").filter(extentElem => extentElem.find("./ExtentType")?.text?.toString() === '09')[0]
      const extentUnit = extentElem?.findtext("./ExtentUnit")?.toString()
      const extentValue = extentElem?.findtext("./ExtentValue")?.toString()
      const bisac = productElement.findtext(".//BASICMainSubject")?.toString()
      const { primaryCategory , secondaryCategory } = this.getCategories(bisac ?? "")
      if (onixVer < 3 && onixVer >= 2) {
        const partialBook: Partial<BookMetadata> = {
          isbn: productElement.findall(".//ProductIdentifier").filter(productId => productId.findtext('ProductIDType')?.toString() === '15')[0].find('IDValue')?.text?.toString(),
          title: productElement.findtext(".//TitleText")?.toString(),
          subtitle: productElement.findtext(".//Subtitle")?.toString(),
          contributors: productElement.findall(".//Contributor").reduce((acc, contributor) => {
            const name = contributor.findtext(".//PersonName")?.toString()
            return name ? [...acc, name] : acc
          }, [] as string[]).join(", "),
          imprint: productElement.findtext(".//ImprintName")?.toString(),
          pubDate: productElement.findtext(".//PublicationDate")?.toString(),
          onSaleDate: productElement.findtext(".//OnSaleDate")?.toString(),
          usPrice: productElement.findall(".//Price").filter(price => price.findtext("./CurrencyCode")?.toString() === "USD")[0]?.findtext("./PriceAmount")?.toString(),
          caPrice: productElement.findall(".//Price").filter(price => price.findtext("./CurrencyCode")?.toString() === "USD")[0]?.findtext("./PriceAmount")?.toString(),
          runtime: this.formatRuntime(extentUnit, extentValue),
          BISAC: productElement.findtext(".//BASICMainSubject")?.toString(),
          language: productElement.findtext(".//LanguageCode")?.toString(),
          primaryCategory,
          secondaryCategories: secondaryCategory,
          customCategory: this.matchSubjectText(productElement.findtext(".//SubjectHeadingText")?.toString() ?? productElement.findall(".//Text").map(text => text.text?.toString()).join("")),
        }
        const book = BookMetadataSchema.safeParse(partialBook)
        if (!book.success) {
          console.error(book.error.flatten().fieldErrors)
          return acc
        }
        additionalBookInfo.set(book.data, { originFile: originFile, onixVer: onixVer })
        return [...acc, book.data]
      } 
      else if (onixVer >= 3 && onixVer < 4) {
        const partialBook: Partial<BookMetadata> = {
          isbn: productElement.findall(".//ProductIdentifier").filter(productId => productId.findtext('ProductIDType')?.toString() === '15')[0].find('IDValue')?.text?.toString(),
          title: (() => {
            const titleText = productElement.findtext(".//TitleText")?.toString()
            const titlePrefix = productElement.findtext(".//TitlePrefix")?.toString()
            const titleWithoutPrefix = productElement.findtext(".//TitleWithoutPrefix")?.toString()
            if (titleText) return titleText
            if (titlePrefix && titlePrefix) return `${titlePrefix} ${titleWithoutPrefix}`
          })(),
          subtitle: productElement.findtext(".//Subtitle")?.toString(),
          contributors: productElement.findall(".//Contributor").reduce((acc, contributor) => {
            const name = contributor.findtext(".//PersonName")?.toString()
            return name ? [...acc, name] : acc
          }, [] as string[]).join(", "),
          imprint: productElement.findtext(".//ImprintName")?.toString(),
          pubDate: productElement.findall(".//PublishingDate").filter(pubDateElem => pubDateElem.findtext("./PublishingDateRole")?.toString() === "01")[0]?.find("./Date")?.text?.toString(),
          onSaleDate: productElement.findall(".//PublishingDate").filter(pubDateElem => pubDateElem.findtext("./PublishingDateRole")?.toString() === "02")[0]?.find("./Date")?.text?.toString(),
          usPrice: productElement.findall(".//Price").filter(price => price.findtext("./CurrencyCode")?.toString() === "USD")[0]?.findtext("./PriceAmount")?.toString(),
          caPrice: productElement.findall(".//Price").filter(price => price.findtext("./CurrencyCode")?.toString() === "USD")[0]?.findtext("./PriceAmount")?.toString(),
          runtime: this.formatRuntime(extentUnit, extentValue),
          BISAC: productElement.findall(".//Subject").filter(subjectElem => subjectElem.find("./MainSubject"))[0].findtext("./SubjectCode")?.toString(),
          language: productElement.findtext(".//LanguageCode")?.toString(),
          primaryCategory,
          secondaryCategories: secondaryCategory,
          customCategory: this.matchSubjectText(productElement.findtext(".//SubjectHeadingText")?.toString() ?? productElement.findall(".//Text").map(text => text.text?.toString()).join("")),
        }
        const book = BookMetadataSchema.safeParse(partialBook)
        if (!book.success) {
          console.error(book.error.flatten())
          return acc
        }
        additionalBookInfo.set(book.data, { originFile: originFile, onixVer: onixVer })
        return [...acc, book.data]
      }
      else {
        return acc
      }
    }, Promise.resolve([] as BookMetadata[]))
    .catch((err: Error) => {
      console.error(err.stack)
      return []
    })
    return parsedMetadata
  }
}