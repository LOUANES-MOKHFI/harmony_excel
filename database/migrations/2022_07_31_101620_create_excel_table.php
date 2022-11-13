<?php

use Illuminate\Database\Migrations\Migration;
use Illuminate\Database\Schema\Blueprint;
use Illuminate\Support\Facades\Schema;

class CreateExcelTable extends Migration
{
    /**
     * Run the migrations.
     *
     * @return void
     */
    public function up()
    {
        /*Schema::create('excel', function (Blueprint $table) {
            $table->id();
            $table->string('contact_date_fiche')->nullable();
            $table->string('pour_centre')->nullable();
            $table->string('date_chargement')->nullable();
            $table->string('contact_qualif1')->nullable();
            $table->string('id_total')->nullable();
            $table->string('accord_montant')->nullable();
            $table->string('contact_qualif2')->nullable();
            $table->string('cas_particulier')->nullable();
            $table->string('pa_montant')->nullable();
            $table->string('pa_frequence')->nullable();
            $table->string('adr1_civilite_abrv')->nullable();
            $table->string('contact_nom')->nullable();
            $table->string('contact_prenom')->nullable();
            $table->string('adr2')->nullable();
            $table->string('adr3')->nullable();
            $table->string('adr4_libelle_voie')->nullable();
            $table->string('adr5')->nullable();
            $table->string('contact_cp')->nullable();
            $table->string('contact_ville')->nullable();
            $table->string('contact_email')->nullable();
            $table->string('contact_tel')->nullable();
            $table->string('contact_tel_port')->nullable();
            $table->string('numero_appeler')->nullable();
            $table->string('new_RAISON_SOCIALE')->nullable();
            $table->string('duree')->nullable();
            $table->string('code_marketing')->nullable();
            $table->string('rf_pro')->nullable();
            $table->string('id_client')->nullable();
            $table->string('envoi_sms')->nullable();
            $table->string('envoi_mail')->nullable();
            $table->string('indice')->nullable();
            $table->string('valid_coordonnees')->nullable();
            $table->string('tel_joint')->nullable();
            $table->string('agent')->nullable();
            $table->string('CMK_S_FIELD_DMC_OUT')->nullable();
            $table->string('Commentaire_call1')->nullable();
            $table->timestamps();
        });*/
    }

    /**
     * Reverse the migrations.
     *
     * @return void
     */
    public function down()
    {
        Schema::dropIfExists('excel');
    }
}
