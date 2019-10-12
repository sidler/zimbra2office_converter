<?php


$inputfile = __DIR__.'/Contacts.csv';
$outputfile = __DIR__.'/o365contacts.csv';

$sourceArray = [];
$rows = 0;
$header = [];
if (($handle = fopen($inputfile, 'r')) !== false) {
    while (($data = fgetcsv($handle)) !== false) {
        if (++$rows === 1) {
            $header = $data;
            continue;
        }

        $mappedRow = [];
        foreach ($data as $index => $value) {
            $mappedRow[$header[$index]] = $value;
        }

        $sourceArray[] = $mappedRow;
    }
    fclose($handle);
}
echo 'Read '.$rows.' rows from excel'.PHP_EOL;


//var_dump($sourceArray);

echo 'Mapping to new structure'.PHP_EOL;


$mapping = [
    'Title'                 => 'Anrede',
    'First Name'            => 'Vorname',
    'Middle Name'           => 'Weitere Vornamen',
    'Last Name'             => 'Nachname',
    'Suffix'                => 'Suffix',
    'Company'               => 'Firma',
    'Department'            => 'Abteilung',
    'Job Title'             => 'Position',
    'Business Street'       => 'Adresse geschäftlich',
    'Business Street 2'     => 'Straße geschäftlich',
    'Business Street 3'     => '',
    'Business City'         => 'Ort geschäftlich',
    'Business State'        => 'Region geschäftlich',
    'Business Postal Code'  => 'Postleitzahl geschäftlich',
    'Business Country/Region' => 'Land/Region geschäftlich',
    'Home Street'           => 'Adresse privat',
    'Home Street 2'         => 'Straße privat',
    'Home Street 3'         => '',
    'Home City'             => 'Ort privat',
    'Home State'            => 'Bundesland/Kanton privat',
    'Home Postal Code'      => 'Postleitzahl privat',
    'Home Country/Region'   => 'Land/Region privat',
    'Other Street'          => 'Weitere Straße',
    'Other Street 2'        => '',
    'Other Street 3'        => '',
    'Other City'            => 'Weiterer Ort',
    'Other State'           => 'Weiteres/r Bundesland/Kanton',
    'Other Postal Code'     => 'Weitere Postleitzahl',
    'Other Country/Region'  => 'Weiteres/e Land/Region',
    'Assistant\'s Phone'    => 'Telefon Assistent',
    'Business Fax'          => 'Fax geschäftlich',
    'Business Phone'        => 'Telefon geschäftlich',
    'Business Phone 2'      => 'Telefon geschäftlich 2',
    'Callback'              => 'Rückmeldung',
    'Car Phone'             => 'Autotelefon',
    'Company Main Phone'    => 'Telefon Firma',
    'Home Fax'              => 'Fax privat',
    'Home Phone'            => 'Telefon privat',
    'Home Phone 2'          => 'Telefon privat 2',
    'ISDN'                  => 'ISDN',
    'Mobile Phone'          => 'Mobiltelefon',
    'Other Fax'             => 'Weiteres Fax',
    'Other Phone'           => 'Weiteres Telefon',
    'Pager'                 => 'Pager',
    'Primary Phone'         => 'Haupttelefon',
    'Radio Phone'           => 'Mobiltelefon 2',
    'TTY/TDD Phone'         => 'Telefon für Hörbehinderte',
    'Telex'                 => 'Telex',
    'Account'               => 'Konto',
    'Anniversary'           => 'Jahrestag',
    'Assistant\'s Name'     => 'Name Assistent',
    'Billing Information'   => 'Abrechnungsinformation',
    'Birthday'              => 'Geburtstag',
    'Business Address PO Box' => 'Postfach geschäftlich',
    'Categories'            => 'Kategorien',
    'Children'              => 'Kinder',
    'Directory Server'      => 'Verzeichnisserver',
    'E-mail Address'        => 'E-Mail-Adresse',
    'E-mail Type'           => 'E-Mail-Typ',
    'E-mail Display Name'   => 'E-Mail: Angezeigter Name',
    'E-mail 2 Address'      => 'E-Mail 2: Adresse',
    'E-mail 2 Type' => 'E-Mail 2: Typ',
    'E-mail 2 Display Name' => 'E-Mail 2: Angezeigter Name',
    'E-mail 3 Address' => 'E-Mail 3: Adresse',
    'E-mail 3 Type' => 'E-Mail 3: Typ',
    'E-mail 3 Display Name' => 'E-Mail 3: Angezeigter Name',
    'Gender' => 'Geschlecht',
    'Government ID Number' => 'Regierungsnr.',
    'Hobby' => 'Hobby',
    'Home Address PO Box' => 'Postfach privat',
    'Initials' => 'Initialen',
    'Internet Free Busy' => 'Internet-Frei/Gebucht',
    'Keywords' => 'Stichwörter',
    'Language' => 'Sprache',
    'Location' => 'Ort',
    'Manager\'s Name' => 'Name des/der Vorgesetzten',
    'Mileage' => 'Reisekilometer',
    'Notes' => 'Notizen',
    'Office Location' => 'Büro',
    'Organizational ID Number' => 'Organisationsnr.',
    'Other Address PO Box' => 'Weiteres Postfach',
    'Priority' => 'Priorität',
    'Private' => 'Privat',
    'Profession' => 'Beruf',
    'Referred By' => 'Empfohlen von',
    'Sensitivity' => 'Vertraulichkeit',
    'Spouse' => 'Partner',
    'User 1' => 'Benutzer 1',
    'User 2' => 'Benutzer 2',
    'User 3' => 'Benutzer 3',
    'User 4' => 'Benutzer 4',
    'Web Page' => 'Webseite'
];

$resultArray = [];
$resultArray[] = array_keys($mapping);

foreach ($sourceArray as $i => $sourceRow) {
    $targetRow = [];
    foreach ($mapping as $newKey => $oldKey) {
        if (empty($oldKey)) {
            $targetRow[] = '';

        } else if (!isset($sourceRow[$oldKey])) {
            echo 'Error in row '.$i.', index '.$oldKey.' not found. Row: '.print_r($sourceRow, true);
            return;

        } else {
            //avoid duplicate data
            $var = $sourceRow[$oldKey];
            if ($oldKey === 'Adresse geschäftlich' && $sourceRow[$oldKey] == $sourceRow['Straße geschäftlich']) {
//                echo 'Removing redundant '.$oldKey.PHP_EOL;
                $var = '';
            }

            if ($oldKey === 'Adresse privat' && $sourceRow[$oldKey] == $sourceRow['Straße privat']) {
//                echo 'Removing redundant '.$oldKey.PHP_EOL;
                $var = '';
            }

            if (in_array($oldKey, ['E-Mail-Adresse', 'E-Mail 2: Adresse', 'E-Mail 3: Adresse'])) {
                if (strpos($var, '"') !== false && strpos($var, '<') !== false) {
                    // we need to trim the masked mail-address
                    $varNew = substr($var, strpos($var, '<')+1, -1);
                    echo 'Replaced mail '.$var.' with '.$varNew.PHP_EOL;
                    $var = $varNew;
                }
            }


            $targetRow[] = $var;

        }
    }
    $resultArray[] = $targetRow;
}

echo 'Writing '.count($resultArray).' rows to new file'.PHP_EOL;

$fp = fopen($outputfile, 'w');

foreach ($resultArray as $fields) {
    fputcsv($fp, $fields);
}

fclose($fp);

echo 'Finished.'.PHP_EOL;

/*






*/
