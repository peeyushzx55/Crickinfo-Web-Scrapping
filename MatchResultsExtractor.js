// node MatchResultsExtractor.js --excel=WorldCup.xlsx --dataDir=Teams --source="https://www.espncricinfo.com/series/icc-cricket-world-cup-2019-1144415/match-results"

let minimist = require('minimist');
let axios = require('axios');
let excel = require('excel4node');
let path = require('path');
let pdf = require('pdf-lib');
let fs = require('fs');
let jsdom = require('jsdom');

let args = minimist(process.argv);

let axiosPromise = axios.get(args.source);
axiosPromise.then(function(response) {
    let html = response.data;
    
    let dom = new jsdom.JSDOM(html);
    let document = dom.window.document;
    
    let matches = [];
    let matchInfoDivs = document.querySelectorAll("div.match-score-block");
    for(let i = 0; i < matchInfoDivs.length; i++) {
        let matchdiv = matchInfoDivs[i];
        let match = {
            t1: "",
            t2: "",
            t1s: "",
            t2s: "",
            result: ""
        };

        let teamParas = matchdiv.querySelectorAll("div.name-detail > p.name");
        match.t1 = teamParas[0].textContent;
        match.t2 = teamParas[1].textContent;

        let scoreSpans = matchdiv.querySelectorAll("div.score-detail > span.score");
        if(scoreSpans.length == 2) {
            match.t1s = scoreSpans[0].textContent;
            match.t2s = scoreSpans[1].textContent;
        }
        else if(scoreSpans.length == 1) {
            match.t1s = scoreSpans[0].textContent;
            match.t2s = "";
        }
        else {
            match.t1s = "";
            match.t2s = "";
        }

        let resultSpans = matchdiv.querySelector("div.status-text > span");
        match.result = resultSpans.textContent;

        matches.push(match);
    }

    let matchesJSON = JSON.stringify(matches);
    fs.writeFileSync("matches.json", matchesJSON, "utf-8");
    
    let teams = [];
    for(let i = 0; i < matches.length; i++) {
        putTeamInTeamsArrayIfMissing(teams, matches[i].t1);
        putTeamInTeamsArrayIfMissing(teams, matches[i].t2);
    }
    for(let i = 0; i < matches.length; i++) {
        putMatchInAppropriateTeam(teams, matches[i].t1, matches[i].t2, matches[i].t1s, matches[i].t2s, matches[i].result);
        putMatchInAppropriateTeam(teams, matches[i].t2, matches[i].t1, matches[i].t2s, matches[i].t1s, matches[i].result);
    }

    let teamsJSON = JSON.stringify(teams);
    fs.writeFileSync("teams.json", teamsJSON, "utf-8");

    createExcelFile(teams, args.excel);
    createFolders(teams, args.dataDir);
})

function createFolders(teams, dataDir) {
    if(fs.existsSync(dataDir) == true) {
        fs.rmdirSync(dataDir, {recursive: true});
    }
    fs.mkdirSync(dataDir);

    for (let i = 0; i < teams.length; i++) {
        let teamFN = path.join(dataDir, teams[i].name);
        fs.mkdirSync(teamFN);

        for (let j = 0; j < teams[i].matches.length; j++) {
            let match = teams[i].matches[j];
            createScoreCard(teamFN, teams[i].name, match);
        }
    }
}

function createScoreCard(fileName, teamName, match) {
    let matchFileName = path.join(fileName, match.vs);

    let originalBytes = fs.readFileSync("Template.pdf");
    let promiseToLoadDoc = pdf.PDFDocument.load(originalBytes);
    promiseToLoadDoc.then(function(pdfDoc) {
        let page = pdfDoc.getPage(0);
        page.drawText(teamName, {
            x: 350,
            y: 400,
            size: 18
        });
        page.drawText(match.vs, {
            x: 350,
            y: 370,
            size: 18
        });
        page.drawText(match.selfScore, {
            x: 350,
            y: 340,
            size: 18
        });
        page.drawText(match.oppScore, {
            x: 350,
            y: 310,
            size: 18
        });
        page.drawText(match.result, {
            x: 150,
            y: 247,
            size: 15
        });

        let promiseToSave = pdfDoc.save();
        promiseToSave.then(function(changedBytes) {
            if(fs.existsSync(matchFileName + ".pdf") == true) {
                fs.writeFileSync(matchFileName + "1.pdf", changedBytes);
            }
            else {
                fs.writeFileSync(matchFileName + ".pdf", changedBytes);
            }
        })
    })
}

function createExcelFile(teams, excelFileName) {
    let wb = new excel.Workbook();
    let style = wb.createStyle({
        fill: {
            type: 'pattern',
            patternType: 'solid',
            fgColor: '#ADD8E6'
        },

        font: {
            bold: true,
            underline: true,
            size: 12,
            shadow: true
        },

        border: {
            left: {
                style: 'medium',
                color: '#000000'
            },

            right: {
                style: 'medium',
                color: '#000000'
            },

            top: {
                style: 'medium',
                color: '#000000'
            },

            bottom: {
                style: 'medium',
                color: '#000000'
            },
        }
    })

    let oppStyle = wb.createStyle({
        fill: {
            type: 'pattern',
            patternType: 'solid',
            fgColor: '#D3D3D3'
        },

        border: {
            left: {
                style: 'medium',
                color: '#000000'
            },

            right: {
                style: 'medium',
                color: '#000000'
            },

            top: {
                style: 'medium',
                color: '#000000'
            },

            bottom: {
                style: 'medium',
                color: '#000000'
            },
        }
    })

    for(let i = 0; i < teams.length; i++) {
        let sheet = wb.addWorksheet(teams[i].name);

        sheet.cell(1, 1).string("Opponent").style(style);
        sheet.cell(1, 2).string("Self Score").style(style);
        sheet.cell(1, 3).string("Opp Score").style(style);
        sheet.cell(1, 4).string("Result").style(style);

        for(let j = 0; j < teams[i].matches.length; j++) {
            
            sheet.cell(2 + j, 1).string(teams[i].matches[j].vs).style(oppStyle);
            sheet.cell(2 + j, 2).string(teams[i].matches[j].selfScore);
            sheet.cell(2 + j, 3).string(teams[i].matches[j].oppScore);
            sheet.cell(2 + j, 4).string(teams[i].matches[j].result);
        }
    }

    wb.write(excelFileName);
}

function putTeamInTeamsArrayIfMissing(teams, teamName) {
    let tidx = -1;
    for (let i = 0; i < teams.length; i++) {
        if (teams[i].name == teamName) {
            tidx = i;
            break;
        }
    }

    if (tidx == -1) {
        teams.push({
            name: teamName,
            matches: []
        })
    }
}

function putMatchInAppropriateTeam(teams, homeTeam, oppTeam, selfScore, oppScore, result) {
    let tidx = -1;
    for (let i = 0; i < teams.length; i++) {
        if (teams[i].name == homeTeam) {
            tidx = i;
            break;
        }
    }

    let team = teams[tidx];
    team.matches.push({
        vs: oppTeam,
        selfScore: selfScore,
        oppScore: oppScore,
        result: result
    })
}