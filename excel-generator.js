const ExcelJS = require('exceljs');

async function generateExcel() {
  const workbook = new ExcelJS.Workbook();

  // Haupt-Sheet für die Sitemap-Struktur
  const sitemapSheet = workbook.addWorksheet('Sitemap Struktur');
  sitemapSheet.columns = [
    { header: 'Domain', key: 'domain' },
    { header: 'Ebene 1', key: 'ebene1' },
    { header: 'Ebene 2', key: 'ebene2' }
  ];

  const domain = 'sandruschka.de';

  // Hauptdienstleistungen
  sitemapSheet.addRow({ domain: domain, ebene1: 'Hauptdienstleistungen', ebene2: 'Graphic Recording & Visualisierung' });
  sitemapSheet.addRow({ domain: domain, ebene1: 'Hauptdienstleistungen', ebene2: '/graphic-recording/' });
  sitemapSheet.addRow({ domain: domain, ebene1: 'Hauptdienstleistungen', ebene2: '/graphic-recording/veranstaltungen/' });
  sitemapSheet.addRow({ domain: domain, ebene1: 'Hauptdienstleistungen', ebene2: '/graphic-recording/workshops/' });
  sitemapSheet.addRow({ domain: domain, ebene1: 'Hauptdienstleistungen', ebene2: '/graphic-recording/buergerbeteiligung/' });

  sitemapSheet.addRow({ domain: domain, ebene1: 'Hauptdienstleistungen', ebene2: 'Wissenschaftskommunikation' });
  sitemapSheet.addRow({ domain: domain, ebene1: 'Hauptdienstleistungen', ebene2: '/wissenschaftskommunikation/' });
  sitemapSheet.addRow({ domain: domain, ebene1: 'Hauptdienstleistungen', ebene2: '/wissenschaftskommunikation/medizin-visualisierung/' });
  sitemapSheet.addRow({ domain: domain, ebene1: 'Hauptdienstleistungen', ebene2: '/wissenschaftskommunikation/wissenschafts-comics/' });
  sitemapSheet.addRow({ domain: domain, ebene1: 'Hauptdienstleistungen', ebene2: '/wissenschaftskommunikation/praesentationen/' });

  sitemapSheet.addRow({ domain: domain, ebene1: 'Hauptdienstleistungen', ebene2: 'Visuelle Strategieentwicklung' });
  sitemapSheet.addRow({ domain: domain, ebene1: 'Hauptdienstleistungen', ebene2: '/strategie-visualisierung/' });
  sitemapSheet.addRow({ domain: domain, ebene1: 'Hauptdienstleistungen', ebene2: '/strategie-visualisierung/prozessvisualisierung/' });
  sitemapSheet.addRow({ domain: domain, ebene1: 'Hauptdienstleistungen', ebene2: '/strategie-visualisierung/zielbilder/' });
  sitemapSheet.addRow({ domain: domain, ebene1: 'Hauptdienstleistungen', ebene2: '/strategie-visualisierung/change-management/' });

  sitemapSheet.addRow({ domain: domain, ebene1: 'Hauptdienstleistungen', ebene2: 'Kreativleistungen & Design' });
  sitemapSheet.addRow({ domain: domain, ebene1: 'Hauptdienstleistungen', ebene2: '/kreativ-design/' });
  sitemapSheet.addRow({ domain: domain, ebene1: 'Hauptdienstleistungen', ebene2: '/kreativ-design/ausstellungsdesign/' });
  sitemapSheet.addRow({ domain: domain, ebene1: 'Hauptdienstleistungen', ebene2: '/kreativ-design/corporate-design/' });
  sitemapSheet.addRow({ domain: domain, ebene1: 'Hauptdienstleistungen', ebene2: '/kreativ-design/illustration/' });

  // Ratgeber-Bereich
  sitemapSheet.addRow({ domain: domain, ebene1: 'Ratgeber-Bereich', ebene2: 'Graphic Recording Wissen' });
  sitemapSheet.addRow({ domain: domain, ebene1: 'Ratgeber-Bereich', ebene2: '/ratgeber/graphic-recording/' });
  sitemapSheet.addRow({ domain: domain, ebene1: 'Ratgeber-Bereich', ebene2: '/ratgeber/graphic-recording/einfuehrung/' });
  sitemapSheet.addRow({ domain: domain, ebene1: 'Ratgeber-Bereich', ebene2: '/ratgeber/graphic-recording/techniken/' });
  sitemapSheet.addRow({ domain: domain, ebene1: 'Ratgeber-Bereich', ebene2: '/ratgeber/graphic-recording/tools/' });

  sitemapSheet.addRow({ domain: domain, ebene1: 'Ratgeber-Bereich', ebene2: 'Wissenschaftskommunikation' });
  sitemapSheet.addRow({ domain: domain, ebene1: 'Ratgeber-Bereich', ebene2: '/ratgeber/wissenschaftskommunikation/' });
  sitemapSheet.addRow({ domain: domain, ebene1: 'Ratgeber-Bereich', ebene2: '/ratgeber/wissenschaftskommunikation/komplexe-themen-vereinfachen/' });
  sitemapSheet.addRow({ domain: domain, ebene1: 'Ratgeber-Bereich', ebene2: '/ratgeber/wissenschaftskommunikation/wissenschaft-gesellschaft/' });

  sitemapSheet.addRow({ domain: domain, ebene1: 'Ratgeber-Bereich', ebene2: 'Visualisierung & Prozesse' });
  sitemapSheet.addRow({ domain: domain, ebene1: 'Ratgeber-Bereich', ebene2: '/ratgeber/visualisierung/' });
  sitemapSheet.addRow({ domain: domain, ebene1: 'Ratgeber-Bereich', ebene2: '/ratgeber/visualisierung/prozessvisualisierung-grundlagen/' });
  sitemapSheet.addRow({ domain: domain, ebene1: 'Ratgeber-Bereich', ebene2: '/ratgeber/visualisierung/team-workshops/' });
  sitemapSheet.addRow({ domain: domain, ebene1: 'Ratgeber-Bereich', ebene2: '/ratgeber/visualisierung/buergerbeteiligung-best-practices/' });

  sitemapSheet.addRow({ domain: domain, ebene1: 'Ratgeber-Bereich', ebene2: 'Digitale Visualisierung' });
  sitemapSheet.addRow({ domain: domain, ebene1: 'Ratgeber-Bereich', ebene2: '/ratgeber/digitale-visualisierung/' });
  sitemapSheet.addRow({ domain: domain, ebene1: 'Ratgeber-Bereich', ebene2: '/ratgeber/digitale-visualisierung/miro-tutorials/' });
  sitemapSheet.addRow({ domain: domain, ebene1: 'Ratgeber-Bereich', ebene2: '/ratgeber/digitale-visualisierung/tablet-zeichnen/' });

  // Standardseiten
  sitemapSheet.addRow({ domain: domain, ebene1: 'Standardseiten', ebene2: '/ueber-uns/' });
  sitemapSheet.addRow({ domain: domain, ebene1: 'Standardseiten', ebene2: '/team/' });
  sitemapSheet.addRow({ domain: domain, ebene1: 'Standardseiten', ebene2: '/referenzen/' });
  sitemapSheet.addRow({ domain: domain, ebene1: 'Standardseiten', ebene2: '/kontakt/' });
  sitemapSheet.addRow({ domain: domain, ebene1: 'Standardseiten', ebene2: '/impressum/' });
  sitemapSheet.addRow({ domain: domain, ebene1: 'Standardseiten', ebene2: '/datenschutz/' });

  // Sheets für Keywords, Title und Meta Descriptions
  const hauptdiensteSheet = workbook.addWorksheet('Hauptdienstleistungen Keywords');
  hauptdiensteSheet.columns = [
    { header: 'Keyword', key: 'keyword' },
    { header: 'Title', key: 'title' },
    { header: 'Meta Description', key: 'description' }
  ];

  const ratgeberSheet = workbook.addWorksheet('Ratgeber-Bereich Keywords');
  ratgeberSheet.columns = [
    { header: 'Keyword', key: 'keyword' },
    { header: 'Title', key: 'title' },
    { header: 'Meta Description', key: 'description' }
  ];

  const standardseitenSheet = workbook.addWorksheet('Standardseiten Keywords');
  standardseitenSheet.columns = [
    { header: 'Keyword', key: 'keyword' },
    { header: 'Title', key: 'title' },
    { header: 'Meta Description', key: 'description' }
  ];

  await workbook.xlsx.writeFile('sitemap-struktur.xlsx');
  console.log('sitemap-struktur.xlsx wurde erfolgreich generiert.');
}

generateExcel().catch(err => {
  console.error('Fehler beim Generieren der Excel-Datei:', err);
});
