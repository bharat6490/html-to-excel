const util = require('util')
const fs = require('fs')
const path = require('path')
const handlebars = require('handlebars');
const conversionFactory = require('html-to-xlsx')
const puppeteer = require('puppeteer')
const chromeEval = require('chrome-page-eval')({ puppeteer })
const writeFileAsync = util.promisify(fs.writeFile)
const excelData =  {
	dataRow:[{
		name: 'Margaret Nguyen11',
		Id: 427311,
		Joined: 'June 3, 2010',
		Canceled: 'n/a',
		balance: 141.5,
	},
	{
		name: 'Edvard Galinski',
		Id: 533175,
		Joined: 'January 13, 2011',
		Canceled: 'April 8, 2017',
		balance: 37.5,
	},
	{
		name: 'Hoshi Nakamura',
		Id: 601942,
		Joined: 'July 23, 2012',
		Canceled: 'n/a',
		balance: 10.0,
	},
	{
		name: 'Margaret Nguyen',
		Id: 427311,
		Joined: 'June 3, 2010',
		Canceled: 'n/a',
		balance: 141.5,
	},
]
	}
  
const conversion = conversionFactory({
  extract: async ({ html, ...restOptions }) => {

	const templatePath = './views/invoice.hbs'
	const dynamicPath = './views/dataconverted.html'

	const filePath = path.join(process.cwd(), templatePath)

	const htmlData = fs.readFileSync(filePath, 'utf8').toString()

	const template = handlebars.compile(htmlData)

	const content = template(excelData)

	var writeStream = fs.createWriteStream(dynamicPath, {

	  flags: 'w',

	})

	writeStream.write(content)

	writeStream.end()



	const reportPath = path.join(process.cwd(), dynamicPath)
		const tmpHtmlPath = path.join(process.cwd(), 'input.html')

    // await writeFileAsync(tmpHtmlPath, html)

    const result = await chromeEval({
      ...restOptions,
      html: reportPath,
      scriptFn: conversionFactory.getScriptFn()
    })

    const tables = Array.isArray(result) ? result : [result]

    return tables.map((table) => ({
      name: table.name,
      getRows: async (rowCb) => {
        table.rows.forEach((row) => {
          rowCb(row)
        })
      },
      rowsCount: table.rows.length
    }))
  }
})

async function run () {
  const stream = await conversion(`<table>
  <tr>
    <th rowspan="2">{{Name}}</th>
    <th rowspan="2">ID</th>
    <th colspan="2">Membership Dates</th>
    <th rowspan="2">Balance</th>
  </tr>
  <tr>
    <th>Joined</th>
    <th>Canceled</th>
  </tr>
  <tr>
    <th>Margaret Nguyen</th>
    <td>427311</td>
    <td><time datetime="2010-06-03">June 3, 2010</time></td>
    <td>n/a</td>
    <td>0.00</td>
  </tr>
  <tr>
    <th>Edvard Galinski</th>
    <td>533175</td>
    <td><time datetime="2011-01-13">January 13, 2011</time></td>
    <td><time datetime="2017-04-08">April 8, 2017</time></td>
    <td>37.00</td>
  </tr>
  <tr>
    <th>Hoshi Nakamura</th>
    <td>601942</td>
    <td><time datetime="2012-07-23">July 23, 2012</time></td>
    <td>n/a</td>
    <td>15.00</td>
  </tr>
</table>`)

  stream.pipe(fs.createWriteStream('output.xlsx'))
}

run()
