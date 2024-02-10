import * as fs from 'fs/promises'
import { Element } from "elementtree"
import * as ET from "elementtree"
import { z } from "zod"

const BookMetadataSchema = z.object({
  isbn: z.string(),
  title: z.string(),
  subtitle: z.optional(z.string()),
  authorsAndNarrators: z.array(z.string()),
  imprint: z.string(),
  pubDate: z.string(),
  onSaleDate: z.string(),
  usPrice: z.string(),
  caPrice: z.optional(z.string()),
  runtime: z.optional(z.number()),
  BISAC: z.string(),
  language: z.string(),
  primaryCategory: z.optional(z.string()),
  secondaryCategory: z.optional(z.string()),
  customCategory: z.optional(z.string())
})
type BookMetadata = z.infer<typeof BookMetadataSchema>

export class MetadataParser {
  private readonly pubDirs = ["SimonSchuster", "Hachette"] as const // directories containing the xml files
  // private readonly pubDirs = ["SimonSchuster"] as const // directories containing the xml files
  private async getRawMetadataList() {
    type RawMetadata = {
      onixVer: number,
      productElement: Element
    }
    const metadataList = await this.pubDirs.reduce(async(accPromise, dir) => {
      const pathList = await fs.readdir(`xml_metadata/${dir}`)
      const acc = await accPromise
      const metadataList = await pathList.reduce(async(accPromise, fileName) => {
        const acc = await accPromise
        const xmlString = await fs.readFile(`xml_metadata/${dir}/${fileName}`, 'utf-8')
        const eTree = ET.parse(xmlString)
        const root = eTree.getroot()
        const product = root.findall('Product')

        if(product.length > 1) {
          const rawMetadataList: RawMetadata[] = product.map((product) => {
            return {
              onixVer: parseInt(root.get('release') ?? '0'), 
              productElement: product
            }
          })
          return [...acc, ...rawMetadataList]
        } else {
          const rawMetadata: RawMetadata = {
            onixVer: parseInt(root.get('release') ?? '2.1'), //single product files assumed Hachette & 2.1
            productElement: product[0]
          }
          return [...acc, rawMetadata]
        }
      }, Promise.resolve([] as RawMetadata[]))
      return [...acc, ...metadataList]
    }, Promise.resolve([] as RawMetadata[]))
    return metadataList
  }

  public async parse(): Promise<BookMetadata[]> {
    const rawMetadataList = await this.getRawMetadataList()
    if (!rawMetadataList) return []
    const parsedMetadata = rawMetadataList.reduce(async (accPromise, { onixVer, productElement }) => {
      const acc = await accPromise
      if (onixVer < 3 && onixVer >= 2) {
      // if (onixVer >= 2) {
        const partialBook: Partial<BookMetadata> = {
          isbn: productElement.findall(".//ProductIdentifier").filter(productId => productId.findtext('ProductIDType')?.toString() === '15')[0].find('IDValue')?.text?.toString(),
          title: productElement.findtext(".//TitleText")?.toString(),
          subtitle: productElement.findtext(".//Subtitle")?.toString(),
          authorsAndNarrators: productElement.findall(".//Contributor").reduce((acc, contributor) => {
            const name = contributor.findtext(".//PersonName")?.toString()
            return name ? [...acc, name] : acc
          }, [] as string[]),
          imprint: productElement.findtext(".//ImprintName")?.toString(),
          pubDate: productElement.findtext(".//PublicationDate")?.toString(),
          onSaleDate: productElement.findtext(".//OnSaleDate")?.toString(),
          usPrice: productElement.findall(".//Price").filter(price => price.findtext("./CurrencyCode")?.toString() === "USD")[0]?.findtext("./PriceAmount")?.toString(),
          caPrice: productElement.findall(".//Price").filter(price => price.findtext("./CurrencyCode")?.toString() === "USD")[0]?.findtext("./PriceAmount")?.toString(),
          runtime: (() => {
            const runtime = parseInt(productElement.findtext(".//ExtentValue")?.toString() || '')
            return isNaN(runtime) ? undefined : runtime
          })(),
          BISAC: productElement.findtext(".//BASICMainSubject")?.toString(),
          language: productElement.findtext(".//LanguageCode")?.toString(),
          // primaryCategory,
          // secondaryCategory,
          // customCategory
        }
        const book = BookMetadataSchema.safeParse(partialBook)
        if (!book.success) {
          console.error(book.error.flatten().fieldErrors)
          return acc
        }
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
          authorsAndNarrators: productElement.findall(".//Contributor").reduce((acc, contributor) => {
            const name = contributor.findtext(".//PersonName")?.toString()
            return name ? [...acc, name] : acc
          }, [] as string[]),
          imprint: productElement.findtext(".//ImprintName")?.toString(),
          pubDate: productElement.findall(".//PublishingDate").filter(pubDateElem => pubDateElem.findtext("./PublishingDateRole")?.toString() === "01")[0]?.find("./Date")?.text?.toString(),
          onSaleDate: productElement.findall(".//PublishingDate").filter(pubDateElem => pubDateElem.findtext("./PublishingDateRole")?.toString() === "02")[0]?.find("./Date")?.text?.toString(),
          usPrice: productElement.findall(".//Price").filter(price => price.findtext("./CurrencyCode")?.toString() === "USD")[0]?.findtext("./PriceAmount")?.toString(),
          caPrice: productElement.findall(".//Price").filter(price => price.findtext("./CurrencyCode")?.toString() === "USD")[0]?.findtext("./PriceAmount")?.toString(),
          runtime: (() => {
            const runtime = parseInt(productElement.findtext(".//ExtentValue")?.toString() || '')
            return isNaN(runtime) ? undefined : runtime
          })(),
          BISAC: productElement.findall(".//Subject").filter(subjectElem => subjectElem.find("./MainSubject"))[0].text?.toString(),
          language: productElement.findtext(".//LanguageCode")?.toString(),
        }
        const book = BookMetadataSchema.safeParse(partialBook)
        if (!book.success) {
          console.error(book.error.flatten().fieldErrors)
          return acc
        }
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