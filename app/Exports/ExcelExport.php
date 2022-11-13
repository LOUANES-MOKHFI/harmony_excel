<?php

namespace App\Exports;

use Maatwebsite\Excel\Concerns\FromCollection;
use Maatwebsite\Excel\Concerns\FromQuery;
use Maatwebsite\Excel\Concerns\Exportable;
use Maatwebsite\Excel\Concerns\WithMapping;
use Maatwebsite\Excel\Concerns\WithHeadings;
use Illuminate\Contracts\Support\Responsable;
use Maatwebsite\Excel\Concerns\WithEvents;
use Maatwebsite\Excel\Events\AfterSheet;
use Maatwebsite\Excel\Excel;
use DB;
class ExcelExport implements WithHeadings, FromCollection, WithEvents, Responsable
{
    /**
    * @return \Illuminate\Support\Collection
    */
    use Exportable;
    public function __construct($table)
    {
        $this->table = $table;
    }
   
    public function headings():array {
        return [
             'contact_date_fiche',
             'pour_centre',
             'date_chargement',
             'contact_qualif1',
             'id_total',
             'accord_montant',
             'contact_qualif2',
             'cas_particulier',
             'pa_montant',
             'pa_frequence',
             'adr1_civilite_abrv',
             'contact_nom',
             'contact_prenom',
             'adr2',
             'adr3',
             'adr4_libelle_voie',
             'adr5',
             'contact_cp',
             'contact_ville',
             'contact_email',
             'contact_tel',
             'contact_tel_port',
             'numero_appeler',
             'new_RAISON_SOCIALE',
             'duree',
             'code_marketing',
             'rf_pro',
             'id_client',
             'envoi_sms',
             'envoi_mail',
             'indice',
             'valid_coordonnees',
             'tel_joint',
             'agent',
             'Acceuil :: TELEPHONE_PORTABLE',
             'contact_email1',
             'CMK_S_FIELD_DMC_OUT',
             'Commentaire_call1',

        ];
    }
    public function collection()
    {
        return DB::table($this->table)->where([['contact_tel_port','<>',''],['Acceuil :: TELEPHONE_PORTABLE','']])
                                    ->orWhereColumn('contact_tel_port','<>','Acceuil :: TELEPHONE_PORTABLE')->get();
    }
    
    public function registerEvents(): array
    {
        return [
            AfterSheet::class=> function(AfterSheet $event) {
  
                $event->sheet->getDelegate()->getStyle('A1:AK1')
                        ->getFill()
                        ->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)
                        ->getStartColor()
                        ->setARGB('DD4B39');
                /*$event->sheet->getDelegate()->getStyle('A2:AK2')
                        ->getFill()
                        ->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)
                        ->getStartColor()
                        ->setARGB('blue');*/
  
            },
        ];
    }
    /*
    public function map($data):array{
            return [
                $data->contact_date_fiche,
                $data->pour_centre,
                $data->date_chargement,
                $data->contact_qualif1,
                $data->id_total,
                $data->accord_montant,
                $data->contact_qualif2,
                $data->cas_particulier,
                $data->pa_montant,
                $data->pa_frequence,
                $data->adr1_civilite_abrv,
                $data->contact_nom,
                $data->contact_prenom,
                $data->adr2,
                $data->adr3,
                $data->adr4_libelle_voie,
                $data->adr5,
                $data->contact_cp,
                $data->contact_ville,
                $data->contact_email,
                $data->contact_tel,
                $data->contact_tel_port,
                $data->numero_appeler,
                $data->new_RAISON_SOCIALE,
                $data->duree,
                $data->code_marketing,
                $data->rf_pro,
                $data->id_client,
                $data->envoi_sms,
                $data->envoi_mail,
                $data->indice,
                $data->valid_coordonnees,
                $data->tel_joint,
                $data->agent,
                $data->CMK_S_FIELD_DMC_OUT,
                $data->Commentaire_call1,
                
            ];
        

    }*/
}
