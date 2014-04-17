#!/usr/bin/perl


sub trim {
     my $string = shift;
     $string =~ s/^\s+//;
     $string =~ s/\s+$//;
     return $string;
}


$FIXHEAD=<<EOF;
<?xml version="1.0" encoding="GBK" standalone="yes"?>
<?mso-application progid="Excel.Sheet"?>
<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:o="urn:schemas-microsoft-com:office:office"
 xmlns:x="urn:schemas-microsoft-com:office:excel"
 xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:html="http://www.w3.org/TR/REC-html40">
 <DocumentProperties xmlns="urn:schemas-microsoft-com:office:office">
  <Created>2006-09-13T11:21:51Z</Created>
  <LastSaved>2006-09-13T11:21:55Z</LastSaved>
  <Version>12.00</Version>
 </DocumentProperties>
 <OfficeDocumentSettings xmlns="urn:schemas-microsoft-com:office:office">
  <RemovePersonalInformation/>
 </OfficeDocumentSettings>
 <ExcelWorkbook xmlns="urn:schemas-microsoft-com:office:excel">
  <WindowHeight>11640</WindowHeight>
  <WindowWidth>19200</WindowWidth>
  <WindowTopX>0</WindowTopX>
  <WindowTopY>90</WindowTopY>
  <ProtectStructure>False</ProtectStructure>
  <ProtectWindows>False</ProtectWindows>
 </ExcelWorkbook>
 <Styles>
  <Style ss:ID="Default" ss:Name="Normal">
   <Alignment ss:Vertical="Center"/>
   <Borders/>
   <Font ss:FontName="宋体" x:CharSet="134" ss:Size="11" ss:Color="#000000"/>
   <Interior/>
   <NumberFormat/>
   <Protection/>
  </Style>
 </Styles>
 <Worksheet ss:Name="Sheet1">
  <Table ss:ExpandedColumnCount="65530" ss:ExpandedRowCount="65530" x:FullColumns="1"
   x:FullRows="1" ss:DefaultColumnWidth="54" ss:DefaultRowHeight="13.5">
EOF



$FIXFOOT=<<EOF;
 </Table>
  <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">
   <PageSetup>
    <Header x:Margin="0.3"/>
    <Footer x:Margin="0.3"/>
    <PageMargins x:Bottom="0.75" x:Left="0.7" x:Right="0.7" x:Top="0.75"/>
   </PageSetup>
   <Unsynced/>
   <Print>
    <ValidPrinterInfo/>
    <PaperSizeIndex>9</PaperSizeIndex>
    <HorizontalResolution>200</HorizontalResolution>
    <VerticalResolution>200</VerticalResolution>
   </Print>
   <Selected/>
   <Panes>
    <Pane>
     <Number>3</Number>
     <ActiveCol>2</ActiveCol>
     <RangeSelection>R1C3:R18C3</RangeSelection>
    </Pane>
   </Panes>
   <ProtectObjects>False</ProtectObjects>
   <ProtectScenarios>False</ProtectScenarios>
  </WorksheetOptions>
 </Worksheet>
</Workbook>

EOF


if($#ARGV!=1){
    print "Usage:expxls.pl {username/password\@sid} {select statement}\n";
    exit(0);
}


sub exeSql{
my ($connstr,$sql)=@_;
my $sqlstr=<<EOF;
sqlplus -s $connstr<<!
set newpage 0;
set termout off;
set echo off;
set feedback off;
set heading on;
set pagesize 50000;
set linesize 9999;
set colsep |+|;
$sql;
exit 
!
EOF

my @rest=`$sqlstr`;
return @rest;
}

# 判断是否是数字
sub checkNumber{    
        return shift =~ /^[+\-]?([1-9]\d*|0)(\.\d+)?([eE][+\-]?([1-9]\d*|0)(\.\d+)?)?$/;
}

@rst=exeSql($ARGV[0],$ARGV[1]);

print $FIXHEAD;


#@rst=exeSql('xjbank/xjbank@query','select * from mainunit where rownum<5;');


delete @rst[1];

foreach my $v(@rst){
		
        my @list=split(/\|\+\|/,$v);
        print '<Row ss:AutoFitHeight="0">',"\n";
            foreach(@list){
                     my $v=trim($_);
                    if(checkNumber($v)){
                        printf(' <Cell><Data ss:Type="Number">%s</Data></Cell>',$v); 
                            
                    }else {
                        printf(' <Cell><Data ss:Type="String">%s</Data></Cell>',$v); 
                    }
					print "\n";
            }
        print '   </Row>',"\n";

}

print $FIXFOOT;
