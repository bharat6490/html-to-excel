# html-to-xlsx
[![NPM Version](http://img.shields.io/npm/v/html-to-xlsx.svg?style=flat-square)](https://npmjs.com/package/html-to-xlsx)
[![License](http://img.shields.io/npm/l/html-to-xlsx.svg?style=flat-square)](http://opensource.org/licenses/MIT)
[![Build Status](https://travis-ci.org/pofider/html-to-xlsx.png?branch=master)](https://travis-ci.org/pofider/html-to-xlsx)

> **node.js html to xlsx transformation**

Transformation only supports html table and several basic style properties. No images or charts are currently supported.

## Usage

```js
const util = require('util')
const fs = require('fs')
const conversionFactory = require('html-to-xlsx')
const puppeteer = require('puppeteer')
const chromeEval = require('chrome-page-eval')({ puppeteer })
const writeFileAsync = util.promisify(fs.writeFile)

const conversion = conversionFactory({
  extract: async ({ html, ...restOptions }) => {
    const tmpHtmlPath = path.join('/path/to/temp', 'input.html')

    await writeFileAsync(tmpHtmlPath, html)

    const result = await chromeEval({
      ...restOptions,
      html: tmpHtmlPath,
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
  const stream = await conversion(`<div class="stl_ stl_02">
        <div class="stl_03">
            <img src="data:image/png;base64,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"
                alt="" class="stl_04" />
        </div>
        <div class="stl_view">
            
           
			<div class="stl_05 stl_06">
				<div class="stl_01" style="left:4.4169em;top:4.5795em;"><span class="stl_07 stl_08 stl_09"
						style="word-spacing:0.0069em;">WIOA - EFFECTIVENESS SERVING EMPLOYERS &nbsp;</span></div>
				<div class="stl_01 stl_10" style="left:7.0335em;top:6.1417em;"><span class="stl_11 stl_08 stl_12">{{STATE_NAME}}
						&nbsp;</span></div>
				<div class="stl_01" style="left:4.3869em;top:6.3308em;"><span class="stl_13 stl_08 stl_14">STATE: &nbsp;</span>
				</div>
				<div class="stl_01" style="left:18.1559em;top:6.3585em;"><span class="stl_13 stl_08 stl_15"
						style="word-spacing:0.009em;">PROGRAM YEAR</span><span class="stl_16 stl_08 stl_15"
						style="word-spacing:0.009em;">:</span><span class="stl_16 stl_08 stl_17" style="word-spacing:-0.1496em;">
						&nbsp;</span></div>
				<div class="stl_01 stl_18" style="left:24.4574em;top:6.2525em;"><span class="stl_19 stl_08 stl_20">{{PROGRAM_YEAR}}
						&nbsp;</span></div>
				<div class="stl_01" style="left:29.5868em;top:6.2404em;"><span class="stl_21 stl_08 stl_22"
						style="word-spacing:-0.0002em;">Certified in WIPS: <span>{{CERTIFIED_IN_WIPS}}</span> &nbsp;</span></div>
				<div class="stl_01" style="left:4.3769em;top:7.7123em;"><span class="stl_23 stl_08 stl_24"
						style="word-spacing:-0.0039em;">PERIOD COVERED &nbsp;</span></div>
				<div class="stl_01" style="left:4.3769em;top:8.8814em;"><span class="stl_25 stl_08 stl_26"
						style="word-spacing:0.0167em;">From ( mm/dd/yyyy ) : </span><span class="stl_27 stl_08 stl_22"
						style="word-spacing:-0.0001em;">7/1/2021 8:00 AM EDT &nbsp;</span></div>
				<div class="stl_01" style="left:21.6815em;top:8.8793em;"><span class="stl_25 stl_08 stl_28"
						style="word-spacing:0.0099em;">To ( mm/dd/yyyy ) : </span><span class="stl_27 stl_08 stl_22"
						style="word-spacing:-0.0001em;">6/30/2022 8:00 AM EDT &nbsp;</span></div>
				<div class="stl_01" style="left:4.3969em;top:10.1023em;"><span class="stl_23 stl_08 stl_29"
						style="word-spacing:-0.0098em;">REPORTING AGENCY: &nbsp;</span></div>
				<div class="stl_01" style="left:4.5067em;top:11.2138em;"><span class="stl_21 stl_08 stl_17"
						style="word-spacing:-0em;">{{REPORTING_AGENCY}} &nbsp;</span></div>
				<div class="stl_01" style="left:4.4169em;top:12.7995em;"><span class="stl_07 stl_08 stl_09"
						style="word-spacing:0.0063em;">EFFECTIVENESS SERVING EMPLOYERS &nbsp;</span></div>
				<div class="stl_01" style="left:13.3569em;top:14.1323em;"><span class="stl_23 stl_08 stl_30"
						style="word-spacing:-0.0035em;">Employer Services &nbsp;</span></div>
				<div class="stl_01" style="left:33.9669em;top:14.1323em;"><span class="stl_23 stl_08 stl_31"
						style="word-spacing:-0.0034em;">Establishment Count &nbsp;</span></div>
				<div class="stl_01" style="left:4.3969em;top:16.7523em;"><span class="stl_32 stl_08 stl_33"
						style="word-spacing:-0.0048em;">Employer Information and Support Services &nbsp;</span></div>
				<div class="stl_01" style="left:4.3969em;top:19.3723em;"><span class="stl_32 stl_08 stl_15"
						style="word-spacing:-0.0012em;">Workforce Recruitment Assistance &nbsp;</span></div>
				<div class="stl_01" style="left:4.3969em;top:21.9923em;"><span class="stl_32 stl_08 stl_34"
						style="word-spacing:-0.0049em;">Engaged in Strategic Planning/Economic Development &nbsp;</span></div>
				<div class="stl_01" style="left:4.3969em;top:24.6123em;"><span class="stl_32 stl_08 stl_35"
						style="word-spacing:-0.0055em;">Accessing Untapped Labor Pools &nbsp;</span></div>
				<div class="stl_01" style="left:4.3969em;top:27.2323em;"><span class="stl_32 stl_08 stl_36"
						style="word-spacing:-0.0074em;">Training Services &nbsp;</span></div>
				<div class="stl_01" style="left:36.691em;top:17.3984em;"><span class="stl_21 stl_08 stl_20">{{INFORMATION_SUPPORT_SERVICES}} &nbsp;</span>
				</div>
				<div class="stl_01" style="left:36.691em;top:20.0184em;"><span class="stl_21 stl_08 stl_20">{{RECRUITMENT_ASSISTANCE}} &nbsp;</span>
				</div>
				<div class="stl_01" style="left:37.323em;top:22.6384em;"><span class="stl_21 stl_08 stl_20">{{STRATEGIC_PLANNING_ECONOMIC_DEVELOPMENT}} &nbsp;</span>
				</div>
				<div class="stl_01" style="left:37.0695em;top:25.2584em;"><span class="stl_21 stl_08 stl_20">{{UNTAPPED_LABOR_POOLS}} &nbsp;</span>
				</div>
				<div class="stl_01" style="left:37.5765em;top:27.8784em;"><span class="stl_21 stl_08 stl_17">{{TRAINING_SERVICES}}</span></div>
				<div class="stl_01" style="left:4.3969em;top:29.8523em;"><span class="stl_32 stl_08 stl_37"
						style="word-spacing:-0.006em;">Incumbent Worker Training Services &nbsp;</span></div>
				<div class="stl_01" style="left:4.3969em;top:32.4723em;"><span class="stl_32 stl_08 stl_34"
						style="word-spacing:-0.0055em;">Rapid Response/Business Downsizing Assistance &nbsp;</span></div>
				<div class="stl_01" style="left:4.3969em;top:35.0923em;"><span class="stl_32 stl_08 stl_35"
						style="word-spacing:-0.0063em;">Planning Layoff Response &nbsp;</span></div>
				<div class="stl_01" style="left:37.5765em;top:30.5052em;"><span class="stl_21 stl_08 stl_17">{{INCUMBENT_WORKER_TRAINING_SERVICES}}</span></div>
				<div class="stl_01" style="left:37.5765em;top:33.1184em;"><span class="stl_21 stl_08 stl_17">{{RAPID_RESPONSE_BUSSINESS_DOWNSIZING}}</span></div>
				<div class="stl_01" style="left:37.5765em;top:35.7316em;"><span class="stl_21 stl_08 stl_17">{{PLANNING_LAYOFFS_RESPONSE}}</span></div>
				<div class="stl_01" style="left:13.5869em;top:37.7123em;"><span class="stl_23 stl_08 stl_15"
						style="word-spacing:-0.0041em;">Pilot Approaches &nbsp;</span></div>
				<div class="stl_01" style="left:41.1269em;top:37.7123em;"><span class="stl_23 stl_08 stl_38">Rate &nbsp;</span>
				</div>
				<div class="stl_01" style="left:31.3569em;top:37.7923em;"><span class="stl_23 stl_08 stl_39">Numerator
						&nbsp;</span></div>
				<div class="stl_01" style="left:30.9369em;top:39.1323em;"><span class="stl_23 stl_08 stl_40">Denominator
						&nbsp;</span></div>
				<div class="stl_01" style="left:32.5658em;top:40.5779em;"><span class="stl_41 stl_08 stl_20">{{RETENTION.NUMERATOR}} &nbsp;</span>
				</div>
				<div class="stl_01" style="left:4.3969em;top:40.3923em;"><span class="stl_32 stl_08 stl_37"
						style="word-spacing:-0.0051em;">Retention with Same Employer in the 2nd and 4th Quarters After &nbsp;</span>
				</div>
				<div class="stl_01" style="left:4.3969em;top:41.5623em;"><span class="stl_32 stl_08 stl_17"
						style="word-spacing:-0.0061em;">Exit Rate &nbsp;</span></div>
				<div class="stl_01" style="left:40.8292em;top:41.1178em;"><span class="stl_42 stl_08 stl_17">{{RETENTION.RATE}} &nbsp;</span>
				</div>
				<div class="stl_01" style="left:40.8224em;top:43.7978em;"><span class="stl_42 stl_08 stl_17">{{PENETRATION.RATE}} &nbsp;</span>
				</div>
				<div class="stl_01" style="left:40.8224em;top:46.4778em;"><span class="stl_42 stl_08 stl_17">{{REPEAT_BUSINESS_CUSTOMERS.RATE}} &nbsp;</span>
				</div>
				<div class="stl_01" style="left:32.5658em;top:41.9179em;"><span class="stl_41 stl_08 stl_20">{{RETENTION.DENOMINATOR}} &nbsp;</span>
				</div>
				<div class="stl_01" style="left:4.3969em;top:43.0723em;"><span class="stl_32 stl_08 stl_20"
						style="word-spacing:-0.0075em;">Employer Penetration Rate &nbsp;</span></div>
				<div class="stl_01" style="left:4.3969em;top:45.7523em;"><span class="stl_32 stl_08 stl_34"
						style="word-spacing:-0.0048em;">Repeat Business Customers Rate &nbsp;</span></div>
				<div class="stl_01" style="left:4.3969em;top:48.4323em;"><span class="stl_32 stl_08 stl_43"
						style="word-spacing:-0.002em;">State Established Measure &nbsp;</span></div>
				<div class="stl_01" style="left:32.5658em;top:43.2579em;"><span class="stl_41 stl_08 stl_20">{{PENETRATION.NUMERATOR}} &nbsp;</span>
				</div>
				<div class="stl_01" style="left:32.3756em;top:44.5979em;"><span class="stl_41 stl_08 stl_20">{{PENETRATION.DENOMINATOR}}
						&nbsp;</span></div>
				<div class="stl_01" style="left:32.5658em;top:45.9379em;"><span class="stl_41 stl_08 stl_20">{{REPEAT_BUSINESS_CUSTOMERS.NUMERATOR}} &nbsp;</span>
				</div>
				<div class="stl_01" style="left:32.3756em;top:47.2779em;"><span class="stl_41 stl_08 stl_20">{{REPEAT_BUSINESS_CUSTOMERS.DENOMINATOR}}
						&nbsp;</span></div>
				<div class="stl_01" style="left:4.4369em;top:51.5666em;"><span class="stl_44 stl_08 stl_45"
						style="word-spacing:-0.0125em;">REPORT CERTIFICATION &nbsp;</span></div>
				<div class="stl_01" style="left:4.3969em;top:53.0523em;"><span class="stl_23 stl_08 stl_46"
						style="word-spacing:-0.0007em;">Report Comments: &nbsp;</span></div>
				<div class="stl_01" style="left:4.5267em;top:54.3458em;"><span class="stl_41 stl_08 stl_12"
						style="word-spacing:0.0002em;">{{REPORT_CERTIFICATION.REPORT_COMMENTS}} &nbsp;</span></div>
				<div class="stl_01" style="left:4.3969em;top:56.2623em;"><span class="stl_23 stl_08 stl_43"
						style="word-spacing:-0.0036em;">Name of Certifying Official/Title: &nbsp;</span></div>
				<div class="stl_01" style="left:19.3269em;top:56.2623em;"><span class="stl_23 stl_08 stl_47"
						style="word-spacing:-0.0033em;">Telephone Number: &nbsp;</span></div>
				<div class="stl_01" style="left:29.4069em;top:56.2623em;"><span class="stl_23 stl_08 stl_43"
						style="word-spacing:-0.0038em;">Email Address: &nbsp;</span></div>
				<div class="stl_01 stl_48" style="left:29.5367em;top:57.8789em;"><a href="mailto:jeanette.pickinpaugh@wyo.gov"
						target="_blank"><span class="stl_27 stl_08 stl_12">{{REPORT_CERTIFICATION.EMAIL}} &nbsp;</span></a></div>
				<div class="stl_01" style="left:19.4567em;top:57.8562em;"><span class="stl_27 stl_08 stl_22"
						style="word-spacing:-0.0005em;">{{REPORT_CERTIFICATION.TELEPHONE_NUMBER}} &nbsp;</span></div>
				<div class="stl_01 stl_49" style="left:4.5267em;top:57.8699em;"><span class="stl_27 stl_08 stl_12"
						style="word-spacing:0.0005em;">{{REPORT_CERTIFICATION.NAME}} &nbsp;</span></div>
			</div>
		</div>
	</div>`)

  stream.pipe(fs.createWriteStream('/path/to/output.xlsx'))
}

run()
```

## Supported properties
- `background-color` - cell background color
- `color` - cell foreground color
- `border-left-style` - as well as positions will be transformed into excel cells borders
- `text-align` - text horizontal align in the excel cell
- `vertical-align` - vertical align in the excel cell
- `width` - the excel column will get the highest width, it can be little bit inaccurate because of pixel to excel points conversion
- `height` - the excel row will get the highest height
- `font-size` - font size
- `colspan` - numeric value that merges current column with columns to the right
- `rowspan` - numeric value that merges current row with rows below.
- `overflow` - the excel cell will have text wrap enabled if this is set to `scroll` or `auto`.

## Constructor options

```js
const conversionFactory = require('html-to-xlsx')
const puppeteer = require('puppeteer')
const chromeEval = require('chrome-page-eval')({ puppeteer })
const conversion = conversionFactory({ /*[constructor options here]*/})
```

- `extract` **function** **[required]** - a function that receives some input (an html file path and a script) and should return some data after been evaluated the html passed. the input that the function receives is:
  ```js
  {
    html: <file path to a html file>,
    scriptFn: <string that contains a javascript function to evaluate in the html>,
    timeout: <time in ms to wait for the function to complete, the function should use this value to abort any execution when the time has passed>,
    /*options passed to `conversion` will be propagated to the input of this function too*/
  }
  ```
- `tmpDir` **string** - the directory path that the module is going to use to save temporary files needed during the conversion. defaults to [`require('os').tmpdir()`](https://nodejs.org/dist/latest-v8.x/docs/api/os.html#os_os_tmpdir)
- `timeout` **number** - time in ms to wait for the conversion to complete, when the timeout is reached the conversion is cancelled. defaults to `10000`

## Conversion options

```js
const fs = require('fs')
const conversionFactory = require('html-to-xlsx')
const puppeteer = require('puppeteer')
const chromeEval = require('chrome-page-eval')({ puppeteer })
const conversion = conversionFactory({
  extract: async ({ html, ...restOptions }) => {
    const tmpHtmlPath = path.join('/path/to/temp', 'input.html')

    await writeFileAsync(tmpHtmlPath, html)

    const result = await chromeEval({
      ...restOptions,
      html: tmpHtmlPath,
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

async function main () {
  const stream = await conversion(/* html */, /* extract options */)
}

main()
```

- `html` **string** - the html source that will be transformed to an xlsx, the html should contain a table element
- `extractOptions` **object** - additional options to pass to the specified `extract` function

## License
See [license](https://github.com/pofider/html-to-xlsx/blob/master/LICENSE)
