#!/usr/bin/env node

var fs   = require('fs');
var path = require('path');
var xls  = require('xlsjs');
var stringify = require('csv-stringify');

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

var dataFlat = Array.prototype.concat(
  data.depenses.reduce(flatten.bind(null, 'depenses'), []),
  data.recettes.reduce(flatten.bind(null, 'recettes'), [])
);

var jsonFlat = JSON.stringify(dataFlat, null, '  ');
fs.writeFileSync( 'budget_flat.json', jsonFlat, 'utf8' );

var dataCSV = dataFlat.map( function (d) {
  
  Object.keys(d.years).forEach(function (k) {
    
    d[k] = d.years[k];

  });

  delete d.years;

  return ['type', 'chapter', 'chapterTitle', 'department', 'departmentTitle', 
          'section', 'sectionTitle', 'article', 'articleTitle', '2013', '2014', 
          '2015', '2016', '2017', '2018' ].map(function (k) {
    return d[k];
  });

});

var csvFlat = stringify(dataCSV, {
  columns: ['type', 'chapter', 'chapterTitle', 'department', 'departmentTitle', 
            'section', 'sectionTitle', 'article', 'articleTitle', '2013', '2014', 
            '2015', '2016', '2017', '2018' ],
  header: true
}, function (err, csv) {
  fs.writeFileSync( 'budget_flat.csv', csv, 'utf8' );
});

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

function flatten (type, memo, parent) {

  if ( !parent.children ) {
    parent.type = type;
    parent.article = parent.number;
    parent.articleTitle = parent.title;
    delete parent.number;
    delete parent.title;
    return memo.concat(parent);
  }

  var children = parent.children
    .reduce(flatten.bind(null, type), [])
    .map(function (d) {
      d[ parent.type ] = parent.number;
      d[ parent.type + 'Title' ] = parent.title;
      return d;
    });

  return memo.concat(children);

}

function parseNumber (str) {

  if ( !str ) return 0;

  return parseFloat( str.replace(/,/g, '') );

}
