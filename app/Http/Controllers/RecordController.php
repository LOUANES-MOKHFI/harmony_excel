<?php

namespace App\Http\Controllers;

use Illuminate\Http\Request;
use Illuminate\Support\Facades\Response;
use DB;
use Illuminate\Support\Facades\Http;
use Illuminate\Support\Facades\Auth;

class RecordController extends Controller
{
    public function index(){
        if(Auth::check()){
            return view('records.index');
        }
        return  redirect()->route("login")->withSuccess('You are not allowed to access');
    }
    public function getAllRecords(Request $request){
      
        try {
            $recordsFile =  $request->recordsFile->move(public_path(), $request->recordsFile->getClientOriginalName());
            $newfile_name_records = $request->recordsFile->getClientOriginalName();
            $newfile_nameRecord = explode('.',$newfile_name_records)[0];
            $spreadsheetHour = \PhpOffice\PhpSpreadsheet\IOFactory::load($newfile_name_records);
            
            $worksheetRecords = $spreadsheetHour->getSheet(0);
            $countLine = $worksheetRecords->getHighestDataRow();

            for($i=2;$i<$countLine;$i++){
                $qualificiation = $worksheetRecords->getCell('M'.$i)->getValue();
                if($qualificiation != null){
                    
                        $statusAppel = $worksheetRecords->getCell('J'.$i)->getValue();
                    
                        if($statusAppel == 'Sortant'){
                            $statusAppel = 'OUT';
                        }elseif($statusAppel == 'Entrant'){
                            $statusAppel = 'IN';
                        }
                        $agentId = $worksheetRecords->getCell('A'.$i)->getValue();
                        $telClient = $worksheetRecords->getCell('B'.$i)->getValue();
                        $dateHoure = $worksheetRecords->getCell('F'.$i)->getValue();
                
                        $date = str_replace('-','',explode(' ',$dateHoure)[0]);
                        $hour = str_replace(':','',explode(' ',$dateHoure)[1]);
                        
                        $target_file = 'https://capitalcorp.comunikcrm.info/capitalcorp_enreg/'.$statusAppel.'-'.$agentId.'-'.$telClient.'-'.$date.'-'.$hour.'.mp3';
                        //dd($target_file);
                       // dd(Http::get('https://capitalcorp.comunikcrm.info/capitalcorp_enreg/OUT-7001-0437280093-20220819-094104.mp3')->successful());
                        $filename = $statusAppel.'-'.$agentId.'-'.$telClient.'-'.$date.'-'.$hour.'.mp3';
                        if(Http::get($target_file)->successful()){
                        $file_headers = @get_headers($target_file);
                    
                        if($request->serverType == 'CALL1'){
                            if($file_headers && $file_headers[0] == 'HTTP/1.1 200 OK') {
                                if($qualificiation == "promesse don avec courrier" || $qualificiation == 'promesse don en ligne'){
                                    file_put_contents(public_path('CALL1/DAM/'.$filename), fopen($target_file, 'r'));
                                }elseif($qualificiation == 'indecis Don' || $qualificiation == 'indecis don_old'){
                                    file_put_contents(public_path('CALL1/Indécis Dons/'.$filename), fopen($target_file, 'r'));
                                }
                            }   
                        }elseif($request->serverType == 'CALL2'){
                            if($file_headers[0] == 'HTTP/1.1 200 OK') {
                                if($qualificiation == "avec preuve" || $qualificiation == 'sans preuve'){
                                    file_put_contents(public_path('CALL2/a déjà renvoyé courrier Fidelis/'.$filename), fopen($target_file, 'r'));
                                }elseif($qualificiation == 'a donne autre assoc' || $qualificiation == 'autre' || $qualificiation == 'jamais donne son accord' || $qualificiation == 'ne se souvient plus' || $qualificiation == 'plus les moyens' || $qualificiation == 'refus conjoint'){
                                    file_put_contents(public_path('CALL2/désistement/'.$filename), fopen($target_file, 'r'));
                                }elseif($qualificiation == 'en differe par donateur' || $qualificiation == 'en differe par agent - CB avec crypto' || $qualificiation == 'en differe par agent - CB sans crypto' || $qualificiation == 'en differe par agent - IBAN' || $qualificiation == 'en differe par donateur - CB' || $qualificiation == 'en differe par donateur - IBAN'){
                                    file_put_contents(public_path('CALL2/don en ligne/'.$filename), fopen($target_file, 'r'));
                                }elseif($qualificiation == 'hors cible' || $qualificiation == 'ne parle pas français' || $qualificiation == 'trop age ou malade'){
                                    file_put_contents(public_path('CALL2/hors cible/'.$filename), fopen($target_file, 'r'));
                                }elseif($qualificiation == 'menace plainte arret relance'){
                                    file_put_contents(public_path('CALL2/menace plainte arret relance/'.$filename), fopen($target_file, 'r'));
                                }elseif($qualificiation == 'autre' || $qualificiation == 'donnera plus tard' || $qualificiation == 'dons autres associations' || $qualificiation == 'entreprise' || $qualificiation == 'pas les moyens' || $qualificiation == 'trop solliciter - ne pas rappeler pendant 6 mois' || $qualificiation == 'vient de donner a cette association'){
                                    file_put_contents(public_path('CALL2/refus de répondre/'.$filename), fopen($target_file, 'r'));
                                }elseif($qualificiation == 'avec courrier - PA' || $qualificiation == 'avec courrier - don avec montant' || $qualificiation == 'avec courrier - indecis don' || $qualificiation == 'sans courrier'){
                                    file_put_contents(public_path('CALL2/va renvoyer/'.$filename), fopen($target_file, 'r'));
                                }/*elseif($qualificiation == 'indecis Don'){
                                    file_put_contents(public_path('CALL2/Pa en ligne/'.$filename), fopen($target_file, 'r'));
                                }*/
                            }   
                        }
                        }
                    
                 }
                
            }
            $data['msg'] = "Les enregistrements sont télècharger avec success";
            $data['status'] = 200;
            return redirect()->back()->with(['success'=>"Les enregistrements sont télècharger avec success"]);
            //return response()->json($data);
        } catch (\Throwable $th) {
            return redirect()->back()->with(['error'=>"erreur de system, veuillez contacter le developpeur s'il vous plait"]);
            $data['msg'] = "erreur de system, veuillez contacter le developpeur s'il vous plait";//$th->getMessage();
            $data['status'] = 500;
            //return response()->json($data);
            
        }
    }
}
