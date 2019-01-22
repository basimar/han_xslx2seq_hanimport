#!/usr/bin/env perl

use strict;
use warnings;
no warnings 'uninitialized';

# Ausführen Script mit "perl xlsx2seq_a226.pl"
# Definition Argumente, mit denen das Script ausgeführt werden muss:
# 1. Argument: Input-Dokument (xlsx)
# 2. Argument: Output (Aleph-Sequential zum laden)
# 3. Argument: Systemnummerbeginn

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

# Aktuelle Zeit auslesen für Feld 583/Leader
use Time::Piece;
my $date583 = localtime->strftime('%d.%m.%Y');
my $date008 = localtime->strftime('%y%m%d');


#Catmandu Importer und Exporter
my $importer = Catmandu::Importer::XLSX->new(file => $ARGV[0]);
my $exporter = Catmandu::Exporter::MARC->new(file => $ARGV[1], type => 'ALEPHSEQ', encoding => 'UTF-8');


#Generierung Systemnummberbeginn
my $sysnum = sprintf("%-9.9d", $ARGV[2]);

my $count = $importer->each(sub {
        #Importer liefert Daten als Hashref ($_[0]), wird dereferenziert und in Hash geladen
        my %hash = %{$_[0]};

        # Datenstruktur: Die Spalten der Tabelle werden im Hash %hash gespeichert. Hash-Key ist jeweils die Spaltenüberschrift
        # Zum Auslesen einer bestimmten Spalte $hash{'Spaltenüberschrift'}

        #Ersetzt falsche Apostrophe und Anführungszeichen in allen Spalten
        for my $value (values %hash) {
             $value =~ s/&apos;/\'/g;
             $value =~ s/‘/\'/g;
             $value =~ s/’/\'/g;
             $value =~ s/”/\"/g;
             $value =~ s/„/\"/g;
             $value =~ s/“/\"/g;
             $value =~ s/&#10;/ /g;
             $value =~ s/&quot;/\"/g;
             $value =~ s/&amp;/&/g;
             $value =~ s/&lt;/</g;
             $value =~ s/&gt;/>/g;
        };

        #Prüft ob in der entsprechenden Excel-Zeile wirklich Daten vorhanden sind, sonst wird die Verarbeitung abgebrochen
        if ($hash{'data'} ne 'n') {

            #Verarbeitung Fussnoten
            my $printer;
        
            if ((defined $hash{'500pr'}) || (defined $hash{'500ja'}) || (defined $hash{'500jh'}) || (defined $hash{'500or'})) {
    
                if ((defined $hash{'500or'}) && (defined $hash{'500pr'})) {
                    $hash{'500or'} = ', ' . $hash{'500or'}
                }

                if ((defined $hash{'500ja'}) && (( defined $hash{'500or'}) || ( defined $hash{'500pr'}))) {
                    $hash{'500ja'} = ', ' . $hash{'500ja'}
                }

                if ((defined $hash{'500jh'}) && (( defined $hash{'500pr'}) || (defined  $hash{'500ja'}) || (defined $hash{'500or'}))) {
                    $hash{'500jh'} = ' (' . $hash{'500jh'} . ')'
                }
            
                $printer = 'Drucker/Entstehungsort der Vorlage: ' . $hash{'500pr'} . $hash{'500or'} . $hash{'500ja'} . $hash{'500jh'};
            }

            my $vorlage;
        
            if ((defined $hash{'500bi'}) && (defined $hash{'500co'}) && (defined $hash{'500bn'}) && (defined $hash{'500si'})) {
                $vorlage = 'Quelle der Vorlage (Musikdruck/Handschrift): [' .  $hash{'500co'} . '] ' . $hash{'500bi'} . ', ' . $hash{'500bn'} . ', ' . $hash{'500si'}
            }

            if (defined $hash{'500au'}) {
                $hash{'500au'} = "Aufgenommene Seiten: " . $hash{'500au'};
            }
         
            # Setze Zeitangaben in Feld 008 auf --------, ausser sie sind in der Tabelle definiert
            unless (defined $hash{'zeit008'} ) {
                $hash{'zeit008'} = '--------';
            }
                        
            #Generiert Data Hash zum Schreiben von Aleph Sequential. Jedes MARC-Feld besteht aus einer Array mit folgenden Elementen
            # Element 0: Feldnummer
            # Element 1: 1. Indikator
            # Element 2: 2. Indikator
            # Element 3: 1. Unterfeldcode
            # Element 4: 1. Unterfeldinhalt
            # Element 5: 2. Unterfeldcode
            # Element 6: 2. Unterfeldinhalt
    
            my $data = {
                _id => $sysnum,
                record => [
                    ['FMT',' ',' ','',$hash{'FMT'}],
                    ['LDR',' ',' ','','-----' . $hash{'LDR'} . '--22-----2u-4500'],
                    ['008',' ',' ','', $date008 . $hash{'806'} . $hash{'zeit008'} . 'xx' . '------------------' . $hash{'835'} . '--'],
                    ['019',' ',' ','a', 'Datenimport Mikrofilmarchiv, Musikwissenschaftliches Seminar Basel' , '5', $date583 . '/A226/eha'],
                    ['040',' ',' ','a','SzZuIDS BS/BE A226'],
                    ['072',' ',' ','a','mu'],
                    ['100',' ',' ','a',$hash{'100'}],
                    ['245',' ',' ','a',$hash{'245a'} . '$$h' . $hash{'245h'}],
                    ['260',' ',' ','a','[S.l.]$$b[s.n.]$$c' . $hash{'260c'}],
                    ['300',' ',' ','a',$hash{'300'}],
                    ['500',' ',' ','a',$printer],
                    ['500',' ',' ','a',$vorlage],
                    ['500',' ',' ','a',$hash{'500'}],
                    ['500',' ',' ','a',$hash{'500au'}],
                    ['510',' ',' ','a',$hash{'510a1'}],
                    ['510',' ',' ','a',$hash{'510a2'}],
                    ['520',' ',' ','a',$hash{'520a'}],
                    ['700',' ',' ','a',$hash{'700'}],
                    ['909','A',' ','a',$hash{'909'}],
                    ['909','A',' ','a','mfa'],
                    ['906',' ',' ','d',$hash{'906d'}],
                    ['907',' ',' ','f',$hash{'907f'}],
                    ['907',' ',' ','e',$hash{'907e'}],
                    # Exemplare werden als Feld 949 geladen. Definitionen für Exemplargenerierung aufgrund Feld 949 finden sich in dsv01/import/tab_hol_item_create

                    ['949',' ',' ','s','A226' . '$$c226MF' . '$$i21' . '$$zMWI MFA ' . $hash{'sig'} . '$$f' . $hash{'ver'} . '$$m' . $hash{'mat'} . '$$nREKATA22615' . '$$tR' ]
                ]
            };
    
            #Lädt Hash in Exporter
            $exporter->add($data);
            # Systemnummer um 1 hochsetzen
            $sysnum += 1;
        };
});

$exporter->commit;
exit;
