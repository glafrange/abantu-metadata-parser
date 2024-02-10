import { MetadataParser } from "./metadata-extraction"

async function main() {
  const mParser = new MetadataParser()
  const result = await mParser.parse()
  // console.log(result.filter(x => !x.subtitle))
  console.log(result)
}

main()