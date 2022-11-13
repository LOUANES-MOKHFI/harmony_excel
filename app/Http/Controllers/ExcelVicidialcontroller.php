<?php

namespace App\Http\Controllers;

use Illuminate\Http\Request;
use Maatwebsite\Excel\Excel;
use App\Http\Controllers\Controller;
use Illuminate\Support\Facades\Auth;
use Spatie\SimpleExcel\SimpleExcelWriter;
use Spatie\SimpleExcel\SimpleExcelReader;
use App\Exports\ExcelExport;
use App\Imports\EditFileExcelExport;
use Illuminate\Support\Facades\Response;
use DB;
use Carbon\Carbon;
class ExcelVicidialcontroller extends Controller
{
    public function index(){
        
        $data = [];
        if(Auth::check()){
            return view('excel.vicidial.index',$data);
        }
        return  redirect()->route("login")->withSuccess('You are not allowed to access');
        
    }

    
     public function StatExcel(Request $request){
       
        
        $date = \DateTime::createFromFormat("Y-m-d", $request->date);
        
        $dayName = $date->format("D");
        $day = $date->format("d");
        
        // 1. Validation du fichier uploadé. Extension ".xlsx" autorisée
        /*$this->validate($request, [
            'fichier' => 'bail|required|file|mimes:xlsx',  /// fichier d'exportation
            'fedilis_file' => 'bail|required|file|mimes:xlsx',  //// fichier fidelis
            'Ce_file' => 'bail|required|file|mimes:xlsx',  //// fichier CE
            'agent_file' => 'bail|required|file|mimes:xlsx',  //// fichier Reporting Agent
            'fichierHour' => 'bail|required|file|mimes:xlsx',  //// fichier Reporting Agent Time
            
            
        ]);*/
        
        
        $info = [];
        if(!$request->fichier || !$request->fedilis_file || !$request->Ce_file || !$request->agent_file || !$request->fichierHour){
            $info['msg'] = "S'il vous plait, veuillez insérer tout les fichiers";
            $info['status'] = 500;
            return response()->json($info);
        }
        $debut = 1;
        if($dayName == 'Mon'){
            $debut = 0;
        }
        elseif($dayName == 'Tue'){
            $debut = 39;
        }
        elseif($dayName == 'Wed'){
            $debut = 78;
        }
        elseif($dayName == 'Thu'){
            $debut = 117;
        }
        elseif($dayName == 'Fri'){
            $debut = 156;
        }
        elseif($dayName == 'Sat'){
            $debut = 195;
        }
        if($debut == 1){
            $info['msg'] = "les Fichier ne contient pas le jour de ".$dayName." ".$date;
            $info['status'] = 500;
            return response()->json($info);
        }

       
        ////exportation File 
        $export = $request->fichier->move(public_path(), $request->fichier->getClientOriginalName());
        $exportname = $request->fichier->getClientOriginalName();
        $newfile_name = explode('.',$exportname)[0];
       // dd($newfile_name);
         ///verify in db exist
         $data['tables'] = DB::select('SHOW TABLES');
         // dd($data['tables']);
          foreach($data['tables'] as $table)
          {   
              if($table->Tables_in_comunik_excel == $newfile_name){
                  $deleteTable = "DROP TABLE $newfile_name";
                  DB::statement($deleteTable);
              }
             // echo $table->Tables_in_db_name;
          }
         
          $createTableSqlString = "CREATE TABLE $newfile_name 
          (
              `id` BIGINT(20) NOT NULL AUTO_INCREMENT , PRIMARY KEY (`id`),
              `contact_date_fiche` varchar(255) NULL,
              `pour_centre` varchar(255) NULL,
              `date_chargement` varchar(255) NULL,
              `contact_qualif1` varchar(255) NULL,
              `id_total` varchar(255) NULL,
              `accord_montant` varchar(255) NULL,
              `contact_qualif2` varchar(255) NULL,
              `cas_particulier` varchar(255) NULL,
              `pa_montant` varchar(255) NULL,
              `pa_frequence` varchar(255) NULL,
              `adr1_civilite_abrv` varchar(255) NULL,
              `contact_nom` varchar(255) NULL,
              `contact_prenom` varchar(255) NULL,
              `adr2` varchar(255) NULL,
              `adr3` varchar(255) NULL,
              `adr4_libelle_voie` varchar(255) NULL,
              `adr5` varchar(255) NULL,
              `contact_cp` varchar(255) NULL,
              `contact_ville` varchar(255) NULL,
              `contact_email` varchar(255) NULL,
              `contact_tel` varchar(255) NULL,
              `contact_tel_port` varchar(255) NULL,
              `numero_appeler` varchar(255) NULL,
              `new_RAISON_SOCIALE` varchar(255) NULL,
              `duree` varchar(255) NULL,
              `code_marketing` varchar(255) NULL,
              `rf_pro` varchar(255) NULL,
              `id_client` varchar(255) NULL,
              `envoi_sms` varchar(255) NULL,
              `envoi_mail` varchar(255) NULL,
              `indice` varchar(255) NULL,
              `valid_coordonnees` varchar(255) NULL,
              `tel_joint` varchar(255) NULL,
              `agent` varchar(255) NULL,
              `Acceuil :: TELEPHONE_PORTABLE` varchar(255) NULL,
              `contact_email1` varchar(255) NULL,
              `CMK_S_FIELD_DMC_OUT` varchar(255) NULL,
              `Commentaire_call1` TEXT NULL,
              `created_at` varchar(255),
              `updated_at` varchar(255)
          )";
         // dd($createTableSqlString);
          DB::statement($createTableSqlString);
          
            
            //  try {
  
            // 3. $reader : L'instance Spatie\SimpleExcel\SimpleExcelReader
            $reader = SimpleExcelReader::create($export);
            
            // On récupère le contenu (les lignes) du fichier
            $rows = $reader->getRows();
            $data = $rows->toArray();
           //dd($data);
            //dd($data);
            // $rows est une Illuminate\Support\LazyCollection
            foreach (array_chunk($data,1000) as $t)  
            {
                $status = DB::table($newfile_name)->insert($t);
            }
           
                

            $qualification = ['don avec montant','promesse don avec courrier','promesse don en ligne','don en ligne','en differe par donateur',
                'en direct par agent - CB avec crypto','en direct par agent - CB sans crypto','en direct par agent - IBAN','indecis Don',
                'en direct par donateur - IBAN','indecis Don','indecis don_old','pa','promesse pa avec courrier','promesse pa en ligne',
                'refus argumente','autre','donnera plus tard','dons autres associations','entreprise','pas les moyens',
                'trop sollicirer - ne pas rapperler pendant 6 mois','vient de donner a cette association'];
            $qual2 = ['promesse don avec courrier','promesse don en ligne'];
            $qual3 = ['en differe par donateur','en direct par agent - CB avec crypto','en direct par agent - CB sans crypto','en direct par agent - IBAN','en direct par donateur - CB','en direct par donateur - IBAN'];
            $basedir = date('d-m-Y');
                if (!file_exists($basedir)) {
                    mkdir($basedir, 0777, true);
                }   
            ////stat des agents
            $fichierHour =  $request->fichierHour->move(public_path(), $request->fichierHour->getClientOriginalName());
            $newfile_name_hour = $request->fichierHour->getClientOriginalName();
            $newfile_nameHour = explode('.',$newfile_name_hour)[0];
            $spreadsheetHour = \PhpOffice\PhpSpreadsheet\IOFactory::load($newfile_name_hour);
            
            $worksheetHour = $spreadsheetHour->getSheet(0);
            $sumPause = 0;
            $sumProd = 0;
            $sumPresence = 0;
            $sumPauseBrief = 0;
            $sumMenu = 0;
            for($i = 2;$i<100;$i++){
            
                $varPauseBrief =$worksheetHour->getCell('M'.$i)->getValue();
                //dd($varPauseBrief);
                if($varPauseBrief != null){
                    $TablepauseBrief = explode(':',$varPauseBrief);
                    $pauseBriefSec = $TablepauseBrief[2];
                    $pauseBriefMin = $TablepauseBrief[1] * 60;
                    $pauseBriefHour = $TablepauseBrief[0] * 3600;
                    $pauseBrief = ($pauseBriefSec+$pauseBriefMin+$pauseBriefHour)/3600;
                    $sumPauseBrief = $sumPauseBrief + $pauseBrief;
                }

                $varPause =$worksheetHour->getCell('J'.$i)->getValue();
                if($varPause != null){
                    $Tablepause = explode(':',$varPause);
                    $pauseSec = $Tablepause[2];
                    $pauseMin = $Tablepause[1] * 60;
                    $pauseHour = $Tablepause[0] * 3600;
                    $pause = ($pauseSec+$pauseMin+$pauseHour)/3600;
                    $sumPause = $sumPause + $pause;
                }
                $varProd = $worksheetHour->getCell('S'.$i)->getValue();
                if($varProd != null){
                    $Tableprod = explode(':',$varProd);
                    $prodSec = $Tableprod[2];
                    $prodMin = $Tableprod[1] * 60;
                    $prodHour = $Tableprod[0] * 3600;
                    $prod = ($prodSec+$prodMin+$prodHour)/3600;
                    $sumProd = $sumProd + $prod;
                }
                $varPresence =$worksheetHour->getCell('T'.$i)->getValue();
                
                if($varPresence != null){
                    $Tablepresence = explode(':',$varPresence);
                    $presenceSec = $Tablepresence[2];
                    $presenceMin = $Tablepresence[1] * 60;
                    $presenceHour = $Tablepresence[0] * 3600;
                    $presence = ($presenceSec+$presenceMin+$presenceHour)/3600;
                    $sumPresence = $sumPresence + $presence;
                }
                //dd($sumPresence);
                $varMenu =$worksheetHour->getCell('R'.$i)->getValue();
               
                if($varMenu != null){
                    $Tablemenu = explode(':',$varMenu);
                    $menuSec = $Tablemenu[2];
                    $menuMin = $Tablemenu[1] * 60;
                    $menuHour = $Tablemenu[0] * 3600;
                    $menu = ($menuSec+$menuMin+$menuHour)/3600;
                    $sumMenu = $sumMenu + $menu;
                }
            }
            //dd($sumPauseBrief);
            $sumPanne = $sumPresence - ($sumPause + $sumProd+$sumMenu);
           
            ////"CE" file
            $Ce_file = $request->Ce_file->move(public_path(), $request->Ce_file->getClientOriginalName());
            $allfile_name = $request->Ce_file->getClientOriginalName();
            $file_name = explode('.',$allfile_name)[0];
            $path = $request->Ce_file->getClientOriginalName();

            $spreadsheetCE = \PhpOffice\PhpSpreadsheet\IOFactory::load($path);
            
            $worksheetCE = $spreadsheetCE->getSheet(0);
            
              $datesCharg = DB::table($newfile_name)->select('date_chargement')->groupBy('date_chargement')->get();
              //$firstLine = DB::table($newfile_name)->find(1);
              $AllLine = DB::table($newfile_name)->get();
              $countAllLine = $AllLine->count();
              //dd($countAllLine);
            foreach($datesCharg as $date_charg){
                $countCharge = DB::table($newfile_name)->where('date_chargement',$date_charg->date_chargement)->count();
                $pourcentage = ((($countCharge*100)/$countAllLine)/100);
                $finalSumProd = $pourcentage*$sumProd;
                $finalSumPause = $pourcentage*$sumPauseBrief;
                $finalSumPanne = $pourcentage*$sumPanne;
                //dd($sumPanne);
                //dd();
                $name_File = 'UNA_PRP_C1_CAP_'.$date_charg->date_chargement;
                
                //$sumHour = $sum / 3600; /// get sum hour
                
                $countCu = DB::table($newfile_name)->where('date_chargement',$date_charg->date_chargement)
                    ->whereIn('contact_qualif1',$qualification)
                    /*->where('contact_qualif1','LIKE','don avec montant')
                    ->orWhere('contact_qualif1','LIKE','promesse don avec courrier')
                    ->orWhere('contact_qualif1','LIKE','promesse don en ligne')////
                    ->orWhere('contact_qualif1','LIKE','don en ligne')
                    ->orWhere('contact_qualif1','LIKE','en differe par donateur')
                    ->orWhere('contact_qualif1','LIKE','en direct par agent - CB avec crypto')
                    ->orWhere('contact_qualif1','LIKE','en direct par agent - CB sans crypto')
                    ->orWhere('contact_qualif1','LIKE','en direct par agent - IBAN')
                    ->orWhere('contact_qualif1','LIKE','en direct par donateur - CB')
                    ->orWhere('contact_qualif1','LIKE','en direct par donateur - IBAN')/////
                    ->orWhere('contact_qualif1','LIKE','indecis Don')
                    ->orWhere('contact_qualif1','LIKE','indecis don_old')
                    ->orWhere('contact_qualif1','LIKE','pa')
                    ->orWhere('contact_qualif1','LIKE','promesse pa avec courrier')
                    ->orWhere('contact_qualif1','LIKE','promesse pa en ligne')
                    ->orWhere('contact_qualif1','LIKE','refus argumente')
                    ->orWhere('contact_qualif1','LIKE','autre')
                    ->orWhere('contact_qualif1','LIKE','donnera plus tard')
                    ->orWhere('contact_qualif1','LIKE','dons autres associations')
                    ->orWhere('contact_qualif1','LIKE','entreprise')
                    ->orWhere('contact_qualif1','LIKE','pas les moyens')
                    ->orWhere('contact_qualif1','LIKE','trop sollicirer - ne pas rapperler pendant 6 mois')
                    ->orWhere('contact_qualif1','LIKE','vient de donner a cette association')*/
                    ->count();  ///// nbr des appele arguementés
                   // dd($countCu);
                   
                $donsPnctl = DB::table($newfile_name)->where('date_chargement',$date_charg->date_chargement)
                                                     ->whereIn('contact_qualif2',$qual2)
                                                     //->where('contact_qualif2','LIKE','promesse don avec courrier')
                                                     //->orWhere('contact_qualif2','LIKE','promesse don en ligne')
                                                     ->count();  ///// nbr des appele arguementés
                $donsEnLigne = DB::table($newfile_name)->where('date_chargement',$date_charg->date_chargement)
                                                       ->whereIn('contact_qualif2',$qual3)
                                                       //->where('contact_qualif2','LIKE','en differe par donateur')
                                                       //->orWhere('contact_qualif2','Like','en direct par donateur - CB')
                                                       ->count();  ///// nbr des appele arguementés
                $line = 6;
                for ($line=6; $line<33 ; $line++) { 
                    //dd($worksheetCE->getCell('A6')->getValue());
                    /*if($worksheetCE->getCell('A'.$line)->getValue() == $name_File){
                        $info['status'] = 500;
                        $info['msg'] = 'Ce fichier est déja uploader';
                        $line = 33;
                    }else{*/
                        if($worksheetCE->getCell('A'.$line)->getValue() == null)
                        {
                            $worksheetCE->getCell('A'.$line)->setValue($name_File);
                    
                            $worksheetCE->getCell('C'.$line)->setValue(round($finalSumPanne, 2));
                            $worksheetCE->getCell('D'.$line)->setValue(round($finalSumPause, 2));
                            $worksheetCE->getCell('E'.$line)->setValue(round($finalSumProd, 2));
                            $worksheetCE->getCell('G'.$line)->setValue($countCu);
                            $worksheetCE->getCell('J'.$line)->setValue($donsPnctl);
                            //$worksheetCE->getCell('J'.$line)->setValue($worksheetCE->getCell('J'.$line)->getValue()+$donsPnctl);
                            $worksheetCE->getCell('K'.$line)->setValue($donsEnLigne);
                            //$worksheetCE->getCell('B2')->setValue('mokhfi');
                            
                            $writerCE = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheetCE, 'Xlsx');
                            
                            $writerCE->save($basedir.DIRECTORY_SEPARATOR.$file_name.'(new).xlsx');
                            $path = $basedir.DIRECTORY_SEPARATOR.$file_name.'(new).xlsx';
                            $info['ce'] = $path;
                            $line = 33;
                        }
                        $info['status'] = 200;
                        $info['msg'] = "Les donneés sont uploader avec succees";//$th->getMessage();
                    //}

                }
            }
            //////Fedilis file
        $fedilis = $request->fedilis_file->move(public_path(), $request->fedilis_file->getClientOriginalName());
        $newfile_name_fedlis = $request->fedilis_file->getClientOriginalName();
        $newfile_nameFidelis = explode('.',$newfile_name_fedlis)[0];
        $spreadsheetFedilis = \PhpOffice\PhpSpreadsheet\IOFactory::load($newfile_name_fedlis);
        $worksheetFedilis = $spreadsheetFedilis->getSheetByName('Asso_'.date('m').'2022'); 
        $headers = ['E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z','AA','AB','AC','AD','AE','AF','AG','AH','AI'];
        
        foreach($headers as $key => $header){
            //dd((date('m') == '08'));
            if($day == $key+1){
                
                $line = 23;
                $array = [];
                for($line = 23; $line<35;$line++){
                    $backgroundColor = $spreadsheetFedilis->getActiveSheet()->getStyle('A'.$line)->getFill()->getStartColor()->getRGB();
                    //array_push($array,$backgroundColor);
                    if($backgroundColor === '5B9BD5'){
                        $keyword = $worksheetFedilis->getCell('A'.$line)->getValue();

                        //dd($keyword);
                        $result = DB::table($newfile_name)
                        ->select('contact_qualif1','contact_qualif2')
                        ->where('contact_qualif1',$keyword)
                        ->orWhere('contact_qualif2',$keyword)
                        ->count();
                        //dd($result);
                        $worksheetFedilis->getCell($header.$line)->setValue($result);
                        //$worksheetFedilis->getCell('E'.$line)->getValue()->setValue($result);
                        
                        //array_push($array, $result);
                    }
                    
                }
                $line = 61;
                $array = [];
                //$worksheetFedilis->getCell('Q62')->setValue('111');
                //$writerFedilis = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheetFedilis, 'Xlsx'); 
                //$writerFedilis->save($basedir.DIRECTORY_SEPARATOR.'aa.xlsx');
                
                for($line = 61; $line<68;$line++){
                    $backgroundColor = $spreadsheetFedilis->getActiveSheet()->getStyle('A'.$line)->getFill()->getStartColor()->getRGB();
                    //array_push($array,$backgroundColor);
                    if($backgroundColor === '5B9BD5'){
                        $keyword = $worksheetFedilis->getCell('A'.$line)->getValue();
                        if($keyword == "Faux numéros :"){$keyword = "faux numero machine";}
                        elseif($keyword == "Faux numéros Système:"){ $keyword = "faux numero";}
                        elseif($keyword == "Hors cible :"){ $keyword = "hors cible";}
                        
                        elseif($keyword == "Refus de repondre"){ $keyword = "refus de repondre";}
                       // dd($keyword);
                       
                        elseif($keyword == "Répondeurs :"){ $keyword = "repondeur";}
                        elseif($keyword == "Injoignables Permanents :"){ $keyword = "injoignable";}

                        //dd($keyword);
                        $result = DB::table($newfile_name)
                        ->select('contact_qualif1','contact_qualif2')
                        ->where('contact_qualif1',$keyword)
                        ->orWhere('contact_qualif2',$keyword)
                        ->count();
                       
                        $worksheetFedilis->getCell($header.$line)->setValue($result);
                        //$worksheetFedilis->getCell('E'.$line)->getValue()->setValue($result);
                        
                        array_push($array, $result);
                    }
                }
                    //dd($array);
                $nbrGSM = DB::table($newfile_name)->select('contact_tel_port','Acceuil :: TELEPHONE_PORTABLE')
                    ->where([['contact_tel_port','<>',''],['Acceuil :: TELEPHONE_PORTABLE','']])
                    ->orWhereColumn('contact_tel_port','<>','Acceuil :: TELEPHONE_PORTABLE')->count();
                    //dd($nbrGSM);
                $nbrMAIL = DB::table($newfile_name)->select('contact_email','contact_email1')
                    ->where([['contact_email','<>',''],['contact_email1','']])
                    ->orWhereColumn('contact_email','<>','contact_email1')->count();
                //dd($nbrMAIL);
                

                $worksheetFedilis->getCell($header.'70')->setValue($nbrGSM);
                $worksheetFedilis->getCell($header.'71')->setValue($nbrMAIL);
                //dd($sumProd);
                $sumMontant = DB::table($newfile_name)->select('accord_montant')->sum('accord_montant');
                $worksheetFedilis->getCell($header.'75')->setValue($sumMontant);

                $worksheetFedilis->getCell($header.'87')->setValue(round($sumProd,2));
                
                $line = 94;
                $array = [];
                for($line = 94; $line<98;$line++){
                    $backgroundColor = $spreadsheetFedilis->getActiveSheet()->getStyle('A'.$line)->getFill()->getStartColor()->getRGB();
                    //array_push($array,$backgroundColor);
                    if($backgroundColor === '5B9BD5'){
                        $keyword = $worksheetFedilis->getCell('A'.$line)->getValue();
                        if($keyword == 'Don en ligne en direct par agent')
                        {
                            $keyword = ['en direct par agent - CB sans crypto','en direct par agent - CB avec crypto','en direct par agent - IBAN'];
                            $result = DB::table($newfile_name)
                            ->select('contact_qualif1','contact_qualif2')
                            ->whereIn('contact_qualif1',$keyword)
                            ->orWhereIn('contact_qualif2',$keyword)
                            ->count();
                        }
                          
                        if($keyword == 'Don en ligne en Direct par donateur'){
                            $keyword = ['en direct par donateur - CB ','en direct par donateur - IBAN'];
                            $result = DB::table($newfile_name)
                            ->select('contact_qualif1','contact_qualif2')
                            ->whereIn('contact_qualif1',$keyword)
                            ->orWhereIn('contact_qualif2',$keyword)
                            ->count();
                        }
                        if($keyword == 'Don en ligne en différé'){
                            $keyword = ['en differe par donateur'];
                            $result = DB::table($newfile_name)
                            ->select('contact_qualif1','contact_qualif2')
                            ->whereIn('contact_qualif1',$keyword)
                            ->orWhereIn('contact_qualif2',$keyword)
                            ->count();
                        }
                        $result = DB::table($newfile_name)
                            ->select('contact_qualif1','contact_qualif2')
                            ->where('contact_qualif1',$keyword)
                            ->orWhere('contact_qualif2',$keyword)
                            ->count();
                        

                        
                        //dd($result);
                        $worksheetFedilis->getCell($header.$line)->setValue($result);
                        //$worksheetFedilis->getCell('E'.$line)->getValue()->setValue($result);
                        
                        //array_push($array, $result);
                    }
                    
                }
                $line = 102;
                $array = [];
                for($line = 102; $line<106;$line++){
                    $backgroundColor = $spreadsheetFedilis->getActiveSheet()->getStyle('A'.$line)->getFill()->getStartColor()->getRGB();
                    //array_push($array,$backgroundColor);
                    if($backgroundColor === '5B9BD5'){
                        $result = DB::table($newfile_name)
                        ->select('contact_qualif2')
                        ->where('contact_qualif2','LIKE',$worksheetFedilis->getCell('A'.$line)->getValue())
                        ->orWhere('contact_qualif1','LIKE',$worksheetFedilis->getCell('A'.$line)->getValue())
                        ->count();
                        //dd($result);
                        $worksheetFedilis->getCell($header.$line)->setValue($result);
                        //$worksheetFedilis->getCell('E'.$line)->getValue()->setValue($result);
                        
                        //array_push($array, $result);
                    }
                    
                }
                
            }
        }
        
        $writerFedilis = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheetFedilis, 'Xlsx'); 
        $writerFedilis->save($basedir.DIRECTORY_SEPARATOR.$newfile_nameFidelis.'(new).xlsx');
        $path1 = $basedir.DIRECTORY_SEPARATOR.$newfile_nameFidelis.'(new).xlsx';
        $info['fedelis'] = $path1;
        ////reporting Agent File

            
            $agent_file = $request->agent_file->move(public_path(), $request->agent_file->getClientOriginalName());
            $agent_file_name = $request->agent_file->getClientOriginalName();
            $new_agent_file_name = explode('.',$agent_file_name)[0];
            
            $spreadsheetAgent = \PhpOffice\PhpSpreadsheet\IOFactory::load($agent_file_name);
           
            //if($day>=01 && $day<07){
            //if($day>=01 && $day<07){
                if($day == '31'){
                    $worksheetAgent = $spreadsheetAgent->getSheetByName('SEM1'); 
                    //dd('01');
                }
                elseif($day>='01' && $day<'06'){
                    $worksheetAgent = $spreadsheetAgent->getSheetByName('SEM1'); 
                    //dd('07');
                }
                elseif($day>'06' && $day<'13'){
                    $worksheetAgent = $spreadsheetAgent->getSheetByName('SEM2'); 
                    //dd('07');
                }
                elseif($day>'13' && $day<'20'){
                    $worksheetAgent = $spreadsheetAgent->getSheetByName('SEM3'); 
                }
                elseif($day>'20' && $day<'27'){
                    $worksheetAgent = $spreadsheetAgent->getSheetByName('SEM4'); 
                }
                elseif($day>'27'){
                    $worksheetAgent = $spreadsheetAgent->getSheetByName('SEM5'); 
                }
                /*elseif($day>'30' && $day<'32'){
                    $worksheetAgent = $spreadsheetAgent->getSheetByName('SEM6'); 
                }*/
           
            $worksheetHour = $spreadsheetHour->getSheet(0);
            $headerAgent = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z'];
            $array = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z'];
            foreach ($headerAgent as $key => $header1) {           
                foreach ($headerAgent as $key => $header2) {
                    $var = [$header1.$header2];
                    array_push($array,$var);
                }
            }
           
       
        for ($x=8; $x < 100; $x++) {
                $agentName =$worksheetAgent->getCell('B'.$x)->getValue();
            //dd($agentName);
            if($agentName != null){
                $sumProd = 0;
                
                for($j=2;$j<100;$j++){
                    $agentNameHour = $worksheetHour->getCell('A'.$j)->getValue();
                    //dd($agentNameHour);
                    if($agentNameHour != null){
                        if($agentName == $agentNameHour){
                            $varProd = $worksheetHour->getCell('S'.$j)->getValue();
                            if($varProd != null){
                                $Tableprod = explode(':',$varProd);
                                $prodSec = $Tableprod[2];
                                $prodMin = $Tableprod[1] * 60;
                                $prodHour = $Tableprod[0] * 3600;
                                $prod = ($prodSec+$prodMin+$prodHour)/3600;
                                $sumProd = $sumProd + $prod;
                                $countCu = DB::table($newfile_name)->where('agent','LIKE',$agentName)
                                            ->whereIn('contact_qualif1',$qualification)
                                            ->count();  ///// nbr des appele arguementés
                                ////Appel argumentés positive
                                $countCuPos = DB::table($newfile_name)
                                                ->where('agent','LIKE',$agentName)
                                                ->where('accord_montant','<>','')
                                                ->count();
                                $countDP = DB::table($newfile_name)
                                                ->where('agent','LIKE',$agentName)
                                                ->whereIn('contact_qualif2',$qual2)
                                                ->count();
                                $DelEnDirect = DB::table($newfile_name)
                                                ->where('agent','LIKE',$agentName)
                                                ->where('contact_qualif2','LIKE','en direct par agent - CB sans crypto')
                                                ->count();
                                $DelEnDifere = DB::table($newfile_name)
                                                ->where('agent','LIKE',$agentName)
                                                ->where('contact_qualif2','LIKE','en differe par donateur')
                                                ->count();
                                $countPEnligne = DB::table($newfile_name)
                                                ->where('agent','LIKE',$agentName)
                                                ->whereIn('contact_qualif2',['indecis Don','indecis don_old'])
                                                ->count();
                                $countRefus = DB::table($newfile_name)
                                                ->where('agent','LIKE',$agentName)
                                                ->where('contact_qualif1','LIKE','refus de repondre')
                                                ->count();
                                
                                $countRAC = DB::table($newfile_name)
                                                ->where('agent','LIKE',$agentName)
                                                ->where('contact_qualif1','LIKE','raccroche')
                                                ->count();
                                $countMontantDon = DB::table($newfile_name)
                                                ->where('agent','LIKE',$agentName)
                                                ->sum('accord_montant');
                                                
                                                
                                if($debut == 0){
                                    //dd($debut);
                                    $worksheetAgent->getCell($array[2+$debut].$x)->setValue(round($sumProd, 2));
                                    //$worksheetAgent->getCell($array[3+$debut].$x)->setValue($countCu);
                                    $worksheetAgent->getCell($array[4+$debut].$x)->setValue($countCuPos);
                                    $worksheetAgent->getCell($array[7+$debut].$x)->setValue($countDP);
                                    $worksheetAgent->getCell($array[8+$debut].$x)->setValue($DelEnDirect);
                                    $worksheetAgent->getCell($array[9+$debut].$x)->setValue($DelEnDifere);
                                    $worksheetAgent->getCell($array[14+$debut].$x)->setValue($countPEnligne);
                                    $worksheetAgent->getCell($array[15+$debut].$x)->setValue($countCu);
                                    $worksheetAgent->getCell($array[16+$debut].$x)->setValue($countRefus);
                                    $worksheetAgent->getCell($array[17+$debut].$x)->setValue($countMontantDon);
                                }else{
                                    //dd($debut);
                                    $worksheetAgent->getCell($array[2+$debut][0].$x)->setValue(round($sumProd, 2));
                                    //$worksheetAgent->getCell($array[3+$debut][0].$x)->setValue($countCu);
                                    $worksheetAgent->getCell($array[4+$debut][0].$x)->setValue($countCuPos);
                                    //$worksheetAgent->getCell($column)->setValue($countCuPos);
                                    $worksheetAgent->getCell($array[7+$debut][0].$x)->setValue($countDP);
                                    $worksheetAgent->getCell($array[8+$debut][0].$x)->setValue($DelEnDirect);
                                    $worksheetAgent->getCell($array[9+$debut][0].$x)->setValue($DelEnDifere);
                                    $worksheetAgent->getCell($array[14+$debut][0].$x)->setValue($countPEnligne);
                                    $worksheetAgent->getCell($array[15+$debut][0].$x)->setValue($countCu);
                                    $worksheetAgent->getCell($array[16+$debut][0].$x)->setValue($countRefus);
                                    $worksheetAgent->getCell($array[17+$debut][0].$x)->setValue($countMontantDon);
                                }
                                
                            }
                        }
                    }
                } 
            }
        }
                
            ///save data in spreadsheet
        $writerAgents = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheetAgent, 'Xlsx');        
        $writerAgents->save($basedir.DIRECTORY_SEPARATOR.$new_agent_file_name.'(new).xlsx');  
        $path2 = $basedir.DIRECTORY_SEPARATOR.$new_agent_file_name.'(new).xlsx';
        $info['agent'] = $path2;
        








           
            
            //return redirect()->back()->with(['success' => "les données sont inserées et le mise à jour est effectuer"]);
            return response()->json($info);
        /*} catch (\Throwable $th) {
            //throw $th;
            $info['msg'] = "erreur de system, le fichier est déja uploader";//$th->getMessage();
            $info['status'] = 500;
            //return redirect()->back()->with(['error' => "erreur de system, le fichier est déja uploader"]);
            return response()->json($info);
        }*/
    }
}
