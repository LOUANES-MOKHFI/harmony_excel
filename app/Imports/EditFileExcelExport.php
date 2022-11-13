<?php

namespace App\Imports;
use Maatwebsite\Excel\Excel;
use PhpOffice\PhpSpreadsheet\Cell\Cell;
use Illuminate\Support\Collection;
use Maatwebsite\Excel\Concerns\ToModel;
use Maatwebsite\Excel\Concerns\WithMappedCells;
use PhpOffice\PhpSpreadsheet\Cell\DataType;
use Maatwebsite\Excel\Concerns\WithCustomValueBinder;
use PhpOffice\PhpSpreadsheet\Cell\DefaultValueBinder;
use Maatwebsite\Excel\Concerns\WithEvents;
use Maatwebsite\Excel\Events\AfterSheet;
class EditFileExcelExport  implements  ToModel, WithMappedCells
{
    
    public function mapping(): array
    {
        return [
            'contact_date_fiche'  => 'A2',
        ];
    }

    public function model(array $row)
    {
        $column = empty($row['contact_date_fiche']) ? $row['contact_date_fiche'] : \PhpOffice\PhpSpreadsheet\IOFactory::load($row['contact_date_fiche']);
        return $column;
    }
    
   
}
