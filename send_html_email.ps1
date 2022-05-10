<#	
	.NOTES
	===========================================================================
     Version:       1.0.0
	 Updated on:   	
	 Created by:   	/u/oRafaelPinheiro
    
	===========================================================================
        RDWebClientManagement Moduile is required
            Install-Module -Name RDWebClientManagement
            https://www.powershellgallery.com/packages/RDWebClientManagement/
        UPDATES
        1.0.1
            /u/oRafaelPinheiro
                - Erro SendEmail
        
	.DESCRIPTION
		Generate a HTML Email about certificate expiration of Windows Remote Access (WRA).
        Send an email when the certificate is about to expire.
    
    .Link
        Original: 
#>

#Certificado utilizado pelo WRA na autenticação do serviço
$Cert_Thumb = $(Get-RDWebClientBrokerCert).Thumbprint

$certs = Get-ChildItem Cert:\LocalMachine\My\$Cert_Thumb `
   | Select @{N='StartDate';E={$_.NotBefore}},
   @{N='DataFim';E={$_.NotAfter}},
   @{N='DiasRestantes';E={($_.NotAfter - (Get-Date)).Days}},
   @{N='Nome';E={$_.DnsNameList}}


$cert_endDate = $certs.DataFim.ToString('dd/MM/yyyy')
$cert_Nome = $($certs.Nome).Unicode
$Dias_Restantes = $certs.DiasRestantes

#Verifica se o certificado está próximo de expirar/vencer
if ($Dias_Restantes -eq 2){

#Generate Header/Style HTML Report
$Header = @"
<style>
    *{
        box-sizing: border-box;
        -webkit-box-sizing: border-box;
        -moz-box-sizing: border-box;
    }

    body{
        font-family: Helvetica;
        -webkit-font-smoothing: antialiased;
    
    }
    h1{
        text-align: center;
        font-size: 18px;
        text-transform: uppercase;
        letter-spacing: 1px;
        color: Black;
        padding: 30px 0;
    }

    h2{
        text-align: center;
        font-size: 18px;
        text-transform: uppercase;
        letter-spacing: 1px;
        color: white;
        padding: 30px 0;
    }


    #fl-table td, #fl-table th {
        text-align: center;
        padding: 8px;
    }

    #fl-table td {
        border-right: 1px solid #f8f8f8;
        font-size: 12px;
    }
</style>
"@

#Generate Body Table HTML Report
$Body = @"
<div>
	<h1>Servidor: $env:computername</h1>
</div>

<div style='margin: 10px 70px 70px; box-shadow: 0px 35px 50px rgba( 0, 0, 0, 0.2 );'>
<table id='fl-table'>

<caption><h3>Informações do Certificado <b style ='text-align: center; color: #3390FF;'> $cert_Nome </b></h3></caption>

<thead>
<tr>
<th style='color: #ffffff; background: #324960;'>Data Expiração</th>
<th style='color: #ffffff;background: #4FC3A1;'>Dias Restantes</th>
<th style='color: #ffffff; background: #324960;'>Certificado</th>
</tr>
</thead>

<tbody>
<tr>
<td>$cert_endDate</td>
<td>$Dias_Restantes</td>
<td><b style ='text-align: center; color: #3390FF;'>$cert_Nome</b></td>
</tr>
<tr>
<td style='text-align: center; background: #F8F8F8;' colspan='3'><h3>Expira em:  <b style='color: #FF1700;'>$Dias_Restantes</b> dias.</h3><td>
</tr>
</tbody>

<thead>
<tr >
<th style='color: #ffffff; background: #324960;' colspan='3'>Procedimentos necessário:</th>
</tr>
</thead>

<tbody>
  <tr>
    <th rowspan='1'>WinACME</th>
    <td colspan='2'>
      <ul style='text-align: left;'>
          <li>Verificar renovação do certificado. Se renovado, ir para IIS</li>
          <li>Verificar thumbnails de Certificados</li>
          <li>Verificar senha do Certificado</li>
       </ul>
    </td>
  </tr>
    
    <tr style='background: #F8F8F8;'>
    <th rowspan='1'>Gerenciador do Serviços de Informações da Internet (IIS)</th>
    <td colspan='2'>
        <ul style='text-align: left;'>
            <li>Verificar certificados do Servidor</li>
            <li>Exportar certificado .pfx</li>
            <li>Sites: Associação Bind e Certificado do Http e Https</li>
        </ul>
    </td>
  </tr>
  
    <tr>
    <th rowspan='1'>Server Manager</th>
    <td colspan="2">
        <ul style='text-align: left;'>
            <li>Serviços de Área de Trabalho Remota - Propriedade de Implantação - Certificados</li>
            <li>Verificar se todos os Serviços de Área de Trabalho Remota estão em execução</li>
        </ul>
    </td>
    </tr>

    <tr style='background: #F8F8F8;'>
    <th rowspan='1'>PowerShell como admin do AD</th>
    <td colspan="2">
        <ul style='text-align: left;'>
            <li>Executar script Certificado_RDWeb.ps1</li>
        </ul>
    </td>
  </tr>
  
</tbody>

<tfoot>
<tr>
<td colspan='3'>(C) 2022 RP SysAdmin $((Get-Date).ToString("dd/MM/yyyy")) </td>
</tr>
</tfoot>

</table>
</div>
"@


#Concatenate Body and Heade and after Convert to HTML
$Report = ConvertTo-HTML -Body $Body -Head $header -Title "Certificado"  | Out-String

#Gera HTML para envio pelo GMAIL
#Generate Report HTML for sending via GMAIL
if ($certs) {
    #smtp server Configuração
    $emailSmtpServer = "smtp.server.com"
    $emailSmtpServerPort = "port"
    $emailSmtpUser = "user_gmail"
    $emailSmtpPass = "cod_pass_gmail"
    
    # recipiente 
    $emailFrom = "from@gmail.com"
    $emailTo = "to@gmail.com"
    $emailcc = "other@other.com"

    # mensagem
    $emailMessage = New-Object System.Net.Mail.MailMessage( $emailFrom , $emailTo )
    $emailMessage.cc.Add($emailcc)
    $emailMessage.Subject = "Certificado Expirando em: $cert_endDate"
    $emailMessage.IsBodyHtml = $true
    $emailMessage.Body = $Report
    
    #cliente 
    $SMTPClient = New-Object System.Net.Mail.SmtpClient( $emailSmtpServer , $emailSmtpServerPort )
    $SMTPClient.EnableSsl = $True
    $SMTPClient.Credentials = New-Object System.Net.NetworkCredential( $emailSmtpUser , $emailSmtpPass );
    $SMTPClient.Send( $emailMessage )
    }
}
