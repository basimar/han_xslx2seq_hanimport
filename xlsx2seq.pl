#!/usr/bin/env perl

use strict;
use warnings;
no warnings 'uninitialized';

die "Argumente: $0 Input-Dokument (xlsx), Output-Dokument, Systemnummerbeginn, \n" unless @ARGV == 3;

# Unicode-Support innerhalb des Perl-Skripts
use utf8;
# Unicode-Support für Output
binmode STDOUT, ":utf8";

# Catmandu-Module
use Catmandu::Importer::XLSX;
use Catmandu::Exporter::MARC;

# Data::Dumper für Debugging
use Data::Dumper;

#Zeit auslesen für Feld 583/Leader
use Time::Piece;
my $date583 = localtime->strftime('%d.%m.%Y');
my $date008 = localtime->strftime('%y%m%d');


#Catmandu Importer und Exporter
my $importer1 = Catmandu::Importer::XLSX->new(file => $ARGV[0]);
my $importer2 = Catmandu::Importer::XLSX->new(file => $ARGV[0]);
my $exporter = Catmandu::Exporter::MARC->new(file => $ARGV[1], type => 'ALEPHSEQ', encoding => 'UTF-8', skip_empty_subfields => 1);


#Generierung Systemnummberbeginn
my $sysnum = sprintf("%-9.9d", $ARGV[2]);

#Generierung %sysnum-Hash für hierarchische Zuordnungen
my %sysnum;
my $count = $importer1->each(sub {
    #Importer liefert Daten als Hashref ($_[0]), wird dereferenziert uns in hash geladen
    my %hash = %{$_[0]};

    foreach (keys %hash) {
        if ($hash{$_} eq "") {
            delete $hash{$_};
        }
    }
	
    #Prüft ob in der entsprechenden Excel-Zeile wirklich Daten vorhanden sind
    if (($hash{'data'} ne 'n') and ($hash{'sys'})) { 
		
        $sysnum = sprintf("%-9.9d", $sysnum);
        #Schreibt Systemnummerpärchen (alt/neu) in %sysnum
       	$sysnum{$hash{'sys'}} = $sysnum;
        print $hash{'sys'}, "\n", $sysnum{$hash{'sys'}}, "\n\n";
        #erhöht neue Systemnummer um 1
        $sysnum = $sysnum + 1;
    };
});

print Dumper(%sysnum);

#Sprachhash für Zuordnung Sprachkürzel und ausgeschriebene Form
my %language = (
    afr => 'Afrikaans',
    alb => 'Albanisch',
    chu => 'Altbulgarisch, Kirchenslawisch',
    grc => 'Altgriechisch',
    san => 'Sanskrit',
    eng => 'Englisch',
    ara => 'Arabisch',
    arc => 'Aramäisch',
    arm => 'Armenisch',
    aze => 'Azeri',
    gez => 'Äthiopisch',
    baq => 'Baskisch',
    bel => 'Weissrussisch',
    ben => 'Bengali',
    bur => 'Burmesisch',
    cze => 'Tschechisch',
    bos => 'Bosnisch',
    bul => 'Bulgarisch',
    roh => 'Rätoromanisch',
    spa => 'Spanisch',
    chi => 'Chinesisch',
    dan => 'Dänisch',
    egy => 'Ägyptisch',
    ger => 'Deutsch',
    gsw => 'Schweizerdeutsch',
    gla => 'Gälisch',
    est => 'Estnisch',
    fin => 'Finnisch',
    dut => 'Niederländisch',
    fre => 'Französisch',
    gle => 'Gälisch',
    geo => 'Georgisch',
    gre => 'Neugriechisch',	
    heb => 'Hebräisch',
    hin => 'Hindi',
    ind => 'Indonesisch',
    ice => 'Isländisch',
    ita => 'Italienisch',
    jpn => 'Japanisch',
    yid => 'Jiddisch',
    khm => 'Khmer',
    kaz => 'Kasachisch',
    kas => 'Kashmiri',
    kir => 'Kirisisch',
    swa => 'Swahili',
    ukr => 'Ukrainisch',
    cop => 'Koptisch',
    kor => 'Koreanisch',
    hrv => 'Kroatisch',
    kur => 'Kurdisch',
    lat => 'Lateinisch',
    lav => 'Lettisch',
    lit => 'Litauisch',
    hun => 'Ungarisch',
    mac => 'Mazedonisch',
    may => 'Malaiisch',
    rum => 'Rumänisch',
    mon => 'Mongolisch',
    per => 'Persisch',
    nor => 'Norwegisch',
    pol => 'Polnisch',
    por => 'Portugiesisch',
    rus => 'Russisch',
    swe => 'Schwedisch',
    srp => 'Serbisch',
    slo => 'Slowakisch',
    slv => 'Slowenisch',
    wen => 'Sorbisch',
    syr => 'Syrisch',
    tgk => 'Tadschikisch',
    tgl => 'Philippinisch',
    tam => 'Tamil',
    tha => 'Siamesisch',
    tur => 'Türkisch',
    tuk => 'Turkmenisch',
    urd => 'Urdu',
    uzb => 'Usbekisch',
    vie => 'Vietnamisch',
    rom => 'Romani'		
);

$count = $importer2->each(sub {
    #Importer liefert Daten als Hashref ($_[0]), wird dereferenziert und in Hash geladen
    my %hash = %{$_[0]};
    #Liest Keys aus dem Hash aus und speichert sie in einer Array
    #my @keys = keys %hash;
    #print "@keys\n";
	
    foreach (keys %hash) {
        if ($hash{$_} eq "") {
            delete $hash{$_};
        }
    }

    #Ersetzt falsche Apostrophe und Anführungszeichen
    for my $value (values %hash) {
        $value =~ s/&apos;/\'/g;
        $value =~ s/‘/\'/g;
        $value =~ s/’/\'/g;
        $value =~ s/”/\"/g;
        $value =~ s/„/\"/g;
        $value =~ s/“/\"/g;
        $value =~ s/&#10;//g;
        $value =~ s/&quot;/\"/g;
        $value =~ s/&amp;/&/g;
        $value =~ s/&lt;/</g;
        $value =~ s/&gt;/>/g;
    };

    #Prüft ob in der entsprechenden Excel-Zeile wirklich Daten vorhanden sind, sonst wird die Verarbeitung abgebrochen
    if (($hash{'data'} ne 'n') && $hash{'sys'}) { 

    #Verarbeitung Materialtyp für LDR: Wenn Materialtyp nicht die Länge 1 hat, wird er auf '-' gesetzt.
    my $mattype = $hash{'LDR1'};
    unless (length $mattype == 1) {$mattype = '-'};

    #Verarbeitung Länder und Sprachcodes in Feld 008 und 546
    #Lädt Sprachcodes in eine Array ein (getrennt durch Komma+Spatium, erstes Arrayelement wird für Feld 008 verwendet
    my @languages = split(', ', $hash{'0081'});
    my $language = $languages[0];
	   	
    #Falls Sprachcode nicht Länge 3 hat, wird er auf 'und' gesetzt
    unless (length $language == 3) {$language = 'und'}; 
		
    #Geht die Liste der Sprachcodes durch, prüft ob die Codes im Sprachhash vorkommen und verkettet die ausführliche Form in der Variable $languange_long		
    my $language_long;

    foreach my $lang (@languages) {
	if (exists($language{$lang})) {$language_long .= $language{$lang} . ', '};
    };
    $language_long =~ s/, ^//g;
 		
    #Wenn nur eine Sprache vorkommt wird lang_041 auf undefiniert gesetzt, damit Feld 041 nicht vergeben wird.
	my $lang_041;
        if (@languages > 1) {$lang_041 = $languages[0]};

        #Bei zweistelligem Ländercode wird ein Bindestrich hinten angefügt, falls danach der Ländercode nicht die Länge 3 hat wird er auf 'xx-' gesetzt
        my $country = $hash{'0082'};		
        if (length $country == 2) {$country .= '-'};
        unless (length $country == 3) {$country = 'xx-'};

        #Verarbeitung Zeitangaben für Feld 260c: Falls in Tabelle ausgefüllt, unverändert übernehmen, falls nicht, Wert aus Feld 593 kopieren
        unless ($hash{'260c'}) {
            $hash{'260c'} = $hash{'593a1'}
        }

        #Verarbeitung der Zeitangaben aus Feld 593 für Codierung in Feld 008: Auslesen des Startjahrs sowie des Endjahres, falls ein Bindestrich vorhanden ist. Ansonsten wird das Endjahr auf '----' gesetzt
        #Falls Bindestrich vorhanden ist, wird die Zeitangabe in Feld 008 mit 'm' codiert, ansonsten mit 's' 
        my $startyear = substr $hash{'593a1'}, 0, 4;
        my $endyear;
        my $timerange;
        my $strich = (index ($hash{'593a1'}, "-")) + 1;
        if ($strich eq 0) {
            $endyear = "----";
            $timerange = 's';
        } else {
            $endyear = substr $hash{'593a1'}, $strich, 4; 	
            $timerange = 'm';
        }

        $startyear = '----' unless length $startyear == 4;
        $endyear = '----' unless length $endyear == 4;


        #Verarbeitung Feld 520a1: Hinzufügen von 'Enthält'
        if ($hash{'520a1'}) {
            $hash{'520a1'} = "Enthält: " . $hash{'520a1'};
        }	
				
        #Verarbeitung Feld 520a22: Hinzufügen von 'Darin'
        if ($hash{'520a2'}) {
            $hash{'520a2'} = "Darin: " . $hash{'520a2'};
        }	 
			
        #Verarbeitung Feld 500a1: Hinzufügen von 'Archivalienart'
        if ($hash{'500a1'}) {
            $hash{'500a1'} = "Archivalienart: " . $hash{'500a1'};
        }	
 
        #Verarbeitung Feld 500a2: Hinzufügen von 'Trägermaterial'
        if ($hash{'500a2'}){
            $hash{'500a2'} = "Trägermaterial: " . $hash{'500a2'};
        }	 
		
        #Verarbeitung Feld 906: Anhand des Inhalts von Feld 906 wird der Indikator entsprechend gesetzt
        my $ind906;
        if ((substr ($hash{'906'}, 0,5) eq "Hochs") || (substr ($hash{'906'}, 0,5) eq "Brief") || (substr ($hash{'906'}, 0,5) eq "Geset")) {$ind906 = "a"};
        if (substr ($hash{'906'}, 0,3) eq "CM ") {$ind906 = "c"};
        if (substr ($hash{'906'}, 0,3) eq "PM ") {$ind906 = "d"};
        if (substr ($hash{'906'}, 0,3) eq "SR ") {$ind906 = "e"};
        if (substr ($hash{'906'}, 0,3) eq "MF ") {$ind906 = "f"};
        if (substr ($hash{'906'}, 0,3) eq "CF ") {$ind906 = "g"};
        if (substr ($hash{'906'}, 0,3) eq "MP ") {$ind906 = "h"};
        if (substr ($hash{'906'}, 0,3) eq "VM ") {$ind906 = "i"};
        if (substr ($hash{'906'}, 0,3) eq "MM ") {$ind906 = "j"};
	
        #Verarbeitung Feld 907: Anhand des Inhalts von Feld 907 wird der Indikator entsprechend gesetzt
        my $ind907;
        if ((substr ($hash{'907'}, 0,5) eq "Hochs") || (substr ($hash{'907'}, 0,5) eq "Brief") || (substr ($hash{'907'}, 0,5) eq "Geset")){$ind907 = "a"};
        if (substr ($hash{'907'}, 0,3) eq "CM ") {$ind907 = "c"};
        if (substr ($hash{'907'}, 0,3) eq "PM ") {$ind907 = "d"};
        if (substr ($hash{'907'}, 0,3) eq "SR ") {$ind907 = "e"};
        if (substr ($hash{'907'}, 0,3) eq "MF ") {$ind907 = "f"};
        if (substr ($hash{'907'}, 0,3) eq "CF ") {$ind907 = "g"};
        if (substr ($hash{'907'}, 0,3) eq "MP ") {$ind907 = "h"};
        if (substr ($hash{'907'}, 0,3) eq "VM ") {$ind907 = "i"};
        if (substr ($hash{'907'}, 0,3) eq "MM ") {$ind907 = "j"};

        my $ind9071;
        if ((substr ($hash{'9071'}, 0,5) eq "Hochs") || (substr ($hash{'9071'}, 0,5) eq "Brief") || (substr ($hash{'9071'}, 0,5) eq "Geset")) {$ind9071 = "a"};
        if (substr ($hash{'9071'}, 0,3) eq "CM ") {$ind9071 = "c"};
        if (substr ($hash{'9071'}, 0,3) eq "PM ") {$ind9071 = "d"};
        if (substr ($hash{'9071'}, 0,3) eq "SR ") {$ind9071 = "e"};
        if (substr ($hash{'9071'}, 0,3) eq "MF ") {$ind9071 = "f"};
        if (substr ($hash{'9071'}, 0,3) eq "CF ") {$ind9071 = "g"};
        if (substr ($hash{'9071'}, 0,3) eq "MP ") {$ind9071 = "h"};
        if (substr ($hash{'9071'}, 0,3) eq "VM ") {$ind9071 = "i"};
        if (substr ($hash{'9071'}, 0,3) eq "MM ") {$ind9071 = "j"};

        my $ind9072;
        if ((substr ($hash{'9072'}, 0,5) eq "Hochs") || (substr ($hash{'9072'}, 0,5) eq "Brief") || (substr ($hash{'9072'}, 0,5) eq "Geset")) {$ind9072 = "a"};
        if (substr ($hash{'9072'}, 0,3) eq "CM ") {$ind9072 = "c"};
        if (substr ($hash{'9072'}, 0,3) eq "PM ") {$ind9072 = "d"};
        if (substr ($hash{'9072'}, 0,3) eq "SR ") {$ind9072 = "e"};
        if (substr ($hash{'9072'}, 0,3) eq "MF ") {$ind9072 = "f"};
        if (substr ($hash{'9072'}, 0,3) eq "CF ") {$ind9072 = "g"};
        if (substr ($hash{'9072'}, 0,3) eq "MP ") {$ind9072 = "h"};
        if (substr ($hash{'9072'}, 0,3) eq "VM ") {$ind9072 = "i"};
        if (substr ($hash{'9072'}, 0,3) eq "MM ") {$ind9072 = "j"};

        #Verknüpfe Signaturen
        my $signatur = $hash{'852p1'} . ' ' . $hash{'852p2'};

        #Prüfung: Ist verlinkte Aufnahme im Systemnummern-Hash (%sysnum) ? Wenn nicht, Link auf ein existierendes Aleph-Katalogisat: Linker-Systemnummer wird nicht angepasst!
	
        unless ($sysnum{$hash{'490w'}}) {
            $sysnum{$hash{'490w'}} = sprintf("%-9.9d", $hash{'490w'})
        };

        #Generiert Data Hash
        my $data = {
            _id => $sysnum{$hash{'sys'}},
       	    record => [
                ['FMT',' ',' ','',$hash{'FMT'}],
                ['LDR',' ',' ','','-----n' . $mattype . 'm--22-----4u-4500'],
                ['008',' ',' ','', $date008 . $timerange . $startyear . $endyear . $country . '-----------------' . $language . '--'],
                ['041',' ',' ','a',$lang_041,'a', $languages[1],'a',$languages[2],'a',$languages[3],'a',$languages[4],'a',$languages[5],'a', $languages[6],'a',$languages[7],'a',$languages[8],'a', $languages[9],'a',$languages[10]],
                ['046',' ',' ','a',$timerange,'c', $startyear, 'e', $endyear],
                ['245',' ',' ','a',$hash{'245a'},'b',$hash{'245b'},'c',$hash{'245c'},'h', $hash{'245h'}],
                ['250',' ',' ','a',$hash{'250a'}],	
                ['260',' ',' ','a',$hash{'260a'},'c',$hash{'260c'}],
                ['300',' ',' ','a',$hash{'300a'},'b',$hash{'300b'},'c',$hash{'300c'}, 'e', $hash{'300e'} ],
                ['340',' ',' ','a',$hash{'340a'}],
                ['351',' ',' ','c',$hash{'351c'}],
                ['355','0',' ','a',$hash{'3550a'}],
                ['355','0',' ','a',$hash{'3550a1'}],
                ['355','0',' ','a',$hash{'3550a2'}],
                ['490',' ',' ','a',$hash{'490a'},'i',$hash{'490i'},'v',$hash{'490v'},'w',$sysnum{$hash{'490w'}}],
                ['500',' ',' ','a',$hash{'500a'}],
                ['500',' ',' ','a',$hash{'500a1'}],
                ['500',' ',' ','a',$hash{'500a2'}],
                ['525',' ',' ','a',$hash{'525a'}],
                ['506',' ',' ','a',$hash{'506a'}],
                ['520',' ',' ','a',$hash{'520a'}],
                ['520',' ',' ','a',$hash{'520a1'}],
                ['520',' ',' ','a',$hash{'520a2'}],
                ['533',' ',' ','n',$hash{'533n'}],
                ['544','1',' ','n',$hash{'5441n'}],
                ['544','0',' ','n',$hash{'5440n'}],
                ['534',' ',' ','n',$hash{'534n'}],
                ['546',' ',' ','a',"$language_long"],
                ['583','0',' ','b','Verzeichnung=Description=Inventaire','c',$date583,'i','Automatisiert nach HAN-Import'],
                ['583','0',' ','b','Verzeichnung=Description=Inventaire','c',$hash{'583c'},'f','ISAD(G) / HAN Katalogisierungsregeln für Archivbestände', 'i', 'Detailliert', 'k', $hash{'583k'}],
                ['852',' ',' ','n',$hash{'852n'},'a',$hash{'852a'},'b',$hash{'852b'},'p',$signatur,'q',$hash{'852q'}, 'z',$hash{'852z'}, 'x',$hash{'852x'}],
                ['856',' ','1','u',$hash{'856u'},'z',$hash{'856z'}],
                ['906',' ',' ',$ind906,$hash{'906'}],
                ['907',' ',' ',$ind907,$hash{'907'}],
                ['907',' ',' ',$ind9071,$hash{'9071'}],
                ['907',' ',' ',$ind9072,$hash{'9072'}],
     	    ],
        };
    #Lädt Hash in Exporter
    $exporter->add($data);
    };
});
$exporter->commit;
exit;
