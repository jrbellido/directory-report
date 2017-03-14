const fs = require('fs')
const path = require('path')

const commander = require('commander')
const upath = require('upath')
const Excel = require('exceljs')

const pjson = require('./package.json')
const availableFormats = ['excel']

commander
    .version('0.0.1')
    .option('-d, --directory [directory]', 'Target directory')
    .option('-f, --format <format>', 'Format (' + availableFormats.join(', ') + ')', new RegExp('^(' + availableFormats.join('|') + ')$', 'i'), availableFormats[0])
    .option('-o, --output [file]', 'Output filename')
    .parse(process.argv)

var resolveTargetDirectory = () => {
    if (typeof commander.directory !== 'undefined')
        return commander.directory
    else
        return path.resolve('./')
}

const targetDirectory = resolveTargetDirectory()

console.log('Target:', targetDirectory)

var walkSync = function(dir, filelist) {
    const files = fs.readdirSync(dir)

    filelist = filelist || []

    files.forEach(function(file) {
        const fullPath = path.join(dir, file)

        if (fs.statSync(fullPath).isDirectory()) {
            filelist = walkSync(fullPath + '/', filelist)
        } else {
            const fileDetails = {
                "file_name": file,
                "file_path": fullPath,
                "file_extension": path.extname(file),
                "normalized_path": upath.normalize(fullPath),
                "file_stats": fs.lstatSync(fullPath)
            }
            filelist.push(fileDetails)
        }
    })

    return filelist
}

const filelist = walkSync(targetDirectory)

const ts = Math.round(new Date().getTime() / 1000)
const fileList = walkSync(targetDirectory)
const format = commander.format
const outputFile = commander.output || 'output_' + ts + '.xlsx'

if (format == 'excel') {
    var workbook = new Excel.Workbook()
    workbook.creator = pjson.name + ' ' + pjson.version

    //var reportSheet = workbook.addWorksheet('Report')
    var listSheet = workbook.addWorksheet('List')

    listSheet.columns = [{
        header: 'Path (normalized)',
        key: 'npath',
        width: 115
    }, {
        header: 'Name',
        key: 'name',
        width: 30
    }, {
        header: 'Extension',
        key: 'ext',
        width: 10
    }, {
        header: 'File size',
        key: 'size',
        width: 10
    }, {
        header: 'Modified',
        key: 'modified',
        width: 20
    }, {
        header: 'Created',
        key: 'created',
        width: 20
    }, {
        header: 'Path',
        key: 'path',
        width: 115,
        outlineLevel: 1
    }, ]

    listSheet.eachRow((row, rowNumber) => {
        row.font = {
            bold: true
        }
    })

    filelist.forEach((file) => {
        var row = {
            "path": file['file_path'],
            "name": file['file_name'],
            "ext": file['file_extension'],
            "size": file['file_stats']['size'],
            "npath": file['normalized_path'],
            "modified": file['file_stats']['mtime'],
            "created": file['file_stats']['ctime']
        }

        listSheet.addRow(row)
    })

    workbook.xlsx.writeFile(outputFile)
}

console.log(fileList[0])
