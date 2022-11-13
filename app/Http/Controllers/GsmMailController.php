<?php

namespace App\Http\Controllers;

use Illuminate\Http\Request;
use Illuminate\Support\Facades\Auth;
use Spatie\SimpleExcel\SimpleExcelWriter;
use Spatie\SimpleExcel\SimpleExcelReader;
use App\Exports\ExcelExport;
use App\Exports\ExcelExportMail;
use DB;
class GsmMailController extends Controller
{
    public function index(){
        
        $data = [];
        if(Auth::check()){
            return view('excel.gsm_mail');
        }
        return  redirect()->route("login")->withSuccess('You are not allowed to access');
        
    }


    public function getAllgsm_mail(Request $request){
        
        $info = [];
        if(!$request->export_file ){
            $info['msg'] = "S'il vous plait, veuillez insérer le fichier d'exportation";
            $info['status'] = 500;
            return response()->json($info);
        }
        ////exportation File 
        $export = $request->export_file->move(public_path(), $request->export_file->getClientOriginalName());
        $exportname = $request->export_file->getClientOriginalName();
        $newfile_name = explode('.',$exportname)[0];
       
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
          DB::statement($createTableSqlString);
          
            
            //  try {
  
            // 3. $reader : L'instance Spatie\SimpleExcel\SimpleExcelReader
            $reader = SimpleExcelReader::create($export);
            
            // On récupère le contenu (les lignes) du fichier
            $rows = $reader->getRows();
            $data = $rows->toArray();
            //dd($data);
            //dd($data);

            foreach (array_chunk($data,1000) as $t)  
            {
                $status = DB::table($newfile_name)->insert($t);
            }
            if($request->type_export == "GSM"){
                return (new ExcelExport($newfile_name))->download('ExportEditGsm.xlsx');
            }elseif($request->type_export == "MAIL"){
                return (new ExcelExportMail($newfile_name))->download('ExportEditMAIL.xlsx');
            }
            


           
          
            
            
          
           
          
        
        

            
           
          
        
        $writerAgents = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheetAgent, 'Xlsx');        
        $writerAgents->save($basedir.DIRECTORY_SEPARATOR.$new_agent_file_name.'(new).xlsx');  
        $path2 = $new_agent_file_name.'(new).xlsx';
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
