#!/usr/bin/env node

var fs   = require('fs');
var path = require('path');
var xls  = require('xlsjs');

var workbook = xls.readFile('budget.xls');

var sheets = workbook.SheetNames.map(function (name) {

  var lines = xls.utils.sheet_to_json( workbook.Sheets[name] );

  return lines.map(function (d) {

    return {
      type: d['Type'] || '',
      num: d['Numéro'] || '',
      id: d['Alias'] || '',
      title: d['Libellé'] || '',
      2013: parseNumber( d['Compte provisoire 2013'] ),
      2014: parseNumber( d['Budget voté 2014'] ),
      2015: parseNumber( d['2015'] ),
      2016: parseNumber( d['2016'] ),
      2017: parseNumber( d['2017'] ),
      2018: parseNumber( d['2018'] )
    }

  });

});

var data = {
  depenses: parseLines( sheets[0] ),
  recettes: parseLines( sheets[1] )
};

var json = JSON.stringify(data, null, '  ');
fs.writeFileSync( 'budget.json', json, 'utf8' );

function parseLines (lines) {

  var activeChapter = null;
  var activeDepartment = null;
  var activeSection = null;

  var data = [];
  lines.forEach( function (d) {

    switch ( d.type ) {

      case 'Chapitre': {

        if ( activeChapter ) {
          data.push(activeChapter);
        }

        activeChapter = {
          type: 'chapter',
          number: d.num,
          title: d.title,
          years: {},
          children: []
        };

        for (var year = 2013; year <= 2018; year++) {
          activeChapter.years[year] = parseFloat( d[year] ) || 0;
        }
        break;

      }

      case 'Département': {

        if ( activeDepartment ) {

          if ( activeChapter ) {
            activeChapter.children.push(activeDepartment);
          } else {
            data.push(activeDepartment);
          }

        }

        activeDepartment = {
          type: 'department',
          number: d.num,
          title: d.title,
          years: {},
          children: []
        };

        for (var year = 2013; year <= 2018; year++) {
          activeDepartment.years[year] = parseFloat( d[year] ) || 0;
        }

        break;

      }

      case 'Section': {

        if ( activeSection ) {

          if ( activeDepartment ) {
            activeDepartment.children.push(activeSection);
          } else {
            data.push(activeSection);
          }

        }

        activeSection = {
          type: 'section',
          number: d.num,
          title: d.title,
          years: {},
          children: []
        };

        for (var year = 2013; year <= 2018; year++) {
          activeSection.years[year] = parseFloat( d[year] ) || 0;
        }

        break;

      }

      case 'Article': {

        var article = {
          type: 'article',
          number: d.num,
          title: d.title,
          years: {}
        };

        for (var year = 2013; year <= 2018; year++) {
          article.years[year] = parseFloat( d[year] ) || 0;
        }

        activeSection.children.push(article);

        break;

      }

      default:
        throw new Error('Unknown type: '+d.type);

    }


  });

  return data;

}


function parseNumber (str) {

  if ( !str ) return 0;

  return parseFloat( str.replace(/,/g, '') );

}
