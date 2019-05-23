clear
Write-Host "`n#####   NOME:`t`t`t" -noNewLine
Write-Host "InsertCadence.ps1" -ForegroundColor "Yellow"
Write-Host "#####   VERSAO:`t`t`t" -noNewLine
Write-Host "1.1" -ForegroundColor "Yellow"
Write-Host "#####   DESCRICAO:`t`t" -noNewLine			
Write-Host "Insere cadencia de passos em atividades de relogios TOMTOM para importar no Strava" -ForegroundColor "Yellow"
Write-Host "#####   DATA DA CRIACAO:`t" -noNewLine	
Write-Host "13/04/2019" -ForegroundColor "Yellow"
Write-Host "#####   ESCRITO POR:`t`t" -noNewLine		
Write-Host "Gilberto Lima de Oliveira" -ForegroundColor "Yellow"
Write-Host "#####   E-MAIL:`t`t`t" -noNewLine				
Write-Host "gilberto.limadeoliveira@gmail.com" -ForegroundColor "Yellow"
Write-Host "#####   Windows:`t`t" -noNewLine				
Write-Host "Windows 8.1" -ForegroundColor "Yellow"
Write-Host "#####   Atividade de ex:`t"	 -noNewLine
Write-Host "https://www.strava.com/activities/2285616271" -ForegroundColor "Yellow"

######################################################################################################
##																									##
##											SECTION 1												##
##									SELECT CSV and GPX FILES										##
##																									##
######################################################################################################

function error(){
	Write-Host "`tFAILED" -ForegroundColor Red
	Write-Host "`tOpcao Invalida"
	Write-Host "`tArquivo" -noNewLine
	Write-Host " CSV ou GPX " -noNewLine -ForegroundColor Red
	Write-Host "nao encontrado"	
	Write-Host "`tFavor acessar a url" -noNewLine
	Write-host " https://mysports.tomtom.com," -ForegroundColor Yellow
	Write-Host "`tclicar na atividade desejada e compartilhar os arquivos .csv e .gpx no mesmo diretorio desse script"
	Write-Host "`tManter o nome dos arquivos no formato run-YYYYMMDDThhmmss.csv e run-YYYYMMDDThhmmss.csv"
	exit 1

}

function escolha(){
	Write-Host "`n"
	for($ss=0;$ss -lt $seleciona.Length; $ss++){
		[int]$option=$ss+1
		Write-Host -NoNewline "$option) " 
		Write-Host -NoNewline $seleciona[$ss].substring(0,4) -ForegroundColor Green
		Write-host -noNewLine "-" -ForegroundColor Green
		Write-Host -NoNewline $seleciona[$ss].substring(4,2) -ForegroundColor Green
		Write-host -noNewLine "-" -ForegroundColor Green
		Write-Host -noNewLine $seleciona[$ss].substring(6,2) -ForegroundColor Green
		Write-host -noNewLine "-" -ForegroundColor Green
		Write-Host $seleciona[$ss].substring(9,6) -ForegroundColor Green	
	}
	$OP=Read-Host -Prompt "`nSeleciona a data da atividade que deseja adicionar a cadencia [1-$option]"
	if ($op -match '[a-z]'){
		error
	}
	[int]$OP=[int]$OP-1
	$filter=$seleciona[$OP]
	return $filter
}

#pesquisando arquivos csv e gpx
$files=Get-ChildItem .\ |Where-Object {$_.Name -match "^run-[0-9]{8}[tT][0-9]{6}.(gpx|csv)"}

#Saindo se nao encontra nenhum
if(-not($files)){
	error
	}
#Removendo duplicados para exibicao para o usuario
#$seleciona=($files.Name).substring(4,8) |Get-Unique
$seleciona=($files.Name).substring(4,15) |Get-Unique
if($seleciona.Count -gt 1){
	$filter=escolha
}
else{
	$filter=$seleciona
}

#$csvFile=Get-ChildItem .\ |Where-Object {$_.Name -match "^run-$filter[tT][0-9]{6}.csv"}
#$gpxFile=Get-ChildItem .\ |Where-Object {$_.Name -match "^run-$filter[tT][0-9]{6}.gpx"}

$csvFile=Get-ChildItem .\ |Where-Object {$_.Name -match "^run-$filter.csv"}
$gpxFile=Get-ChildItem .\ |Where-Object {$_.Name -match "^run-$filter.gpx"}
if(-not($csvFile)){
	error
}
if(-not($gpxFile)){
	error
}

Write-Host "Arquivos escolhidos: "
Write-Host "`t" $csvFile.Name
Write-Host "`t" $gpxFile.Name
$newgpx="New_"+$gpxFile.name

[char]$OP=Read-Host -Prompt "Confirma? [S/N]"
if($OP -eq "n"){
	exit 1
	Write-Host "Saindo"
}
######################################################################################################
##																									##
##											SECTION 2												##
##						CREATE and INSERT NEW ELEMENT XML WITH RIGHT VALUE							##
##																									##
######################################################################################################

$currentDir=Get-Location
$dirDest="$currentDir\old"

Get-ChildItem .\ |Where-Object {$_.name -match 'New_run-[0-9]{8}T[0-9]{6}.gpx'} |Move-Item -Destination $dirDest -Force

Write-Host "`nTrabalhando nos arquivos," -noNewLine
Write-Host " aguarde!" -ForegroundColor Yellow

#Import GPX
[xml]$gpx=Get-Content $gpxFile
#Impcort csv
$csv=Import-csv $csvFile
sleep 2
#Trata xml namespace 
$gpxtpx=New-Object System.Xml.XmlNamespaceManager($gpx.NameTable)
$gpxtpx.AddNamespace("gpxtpx",$gpx.gpx.GetNamespaceOfPrefix("gpxtpx") 	)
#Seta namespace para pesquisa
$namespace=@{gpx="http://www.topografix.com/GPX/1/1"}
#Seta total de linhas para exicibiÃ§Ã£o
#[int]$latLongTotal=$csv.count
$latLongTotal=[timespan]::fromseconds($csv.count-1)
$latLongTotal="{0:HH:mm:ss}" -f ([datetime]$latLongTotal.Ticks)

[int]$count="0"
$cadmedia=@()
foreach ($movimentacion in $csv){
	#Seta numero da linha atual
	$count=$count+1
	#Define variaveis para latitude, longitude e cadencia para a linha atual
	$lat=$movimentacion.lat
	$long=$movimentacion.long
	$cad=$movimentacion.cycles
	#Testa se os valores do csv nÃ£o estÃ£o vazios
	if (($lat) -and ($long)){
		#Pesquisa no arquivo gpx a lat e long da linha atual (anos trabalhando com FIM2010 R2 e muito xPath)
		$element=Select-Xml -Xml $gpx -XPath "//gpx:trkpt[@lat=$lat and @lon=$long]" -Namespace $namespace
		if($element){
			#Se a lat e long atual exite no csv, comeca a tratar a cadencia
			#Converte a cadencia do csv para formato do strava
			$cadmedia+=$cad
			#Nos primeiros 60s Ã© considerado a media que tem, exemplo: para o segundo 10 soma toda a cadencia e divide por 10, depois multiplica por 30
			if ($cadmedia.Length -lt 61){
			#Nesse IF assume o que estÃ¡ no csv da Tomtom, basicamente a cadencia por seg multiplicada por 30
				[int]$cadf=(($cadmedia |Measure-Object -Sum).sum/$cadmedia.Length)
				$cadf=[int]$cadf*30
			}

			#Quando tem 60 segundos, passa a somar 60 segundos e depois dividir por 2, porque no gpx deve ir a metade dos passos (formato do schema)
			if($cadmedia.Length -gt 60){
			#Nesse ponto Ã© usado a media dos ultimos 60 segundos
				[int]$arrayControl=($cadmedia.Length - 61)
				$cadmedia[$arrayControl]=$null
				[int]$cadf=(($cadmedia |Measure-Object -Sum).sum/2)
			
			}
			#soma cadmedia
			#($cadmedia |Measure-Object -Sum).sum
			
			#Cria o elementp xml para inserir a cadencia
			$cadence=$gpx.CreateElement("gpxtpx:cad", $gpxtpx.LookupNamespace("gpxtpx"))
			
			
			#Insere o elemento no lugar certo na arvore xml
			if($element.Node.extensions.TrackPointExtension){
				$element.Node.extensions.TrackPointExtension.AppendChild($cadence) > $null
			}
			#seta o valor atual para inserir no xml
			$value=$gpx.CreateTextNode($cadf)
			#insere o valor
			$cadence.AppendChild($value) > $null
			#mensagem para o usuario
			$tempoAtv=[timespan]::fromseconds($count)
			$tempoAtv="{0:HH:mm:ss}" -f ([datetime]$tempoAtv.Ticks)
			
			
			#Write-host "Tempo: " -noNewLine
			Write-host "$tempoAtv" -ForegroundColor Red -noNewLine
			#Write-host "Seg: " -noNewLine
			#Write-host "$count" -ForegroundColor Red -noNewLine
			Write-host " de " -noNewLine
			Write-host "$latLongTotal" -ForegroundColor Red -noNewLine
			Write-host " Lat: " -noNewLine
			Write-host "$lat" -ForegroundColor Green -noNewLine
			Write-host " Long: " -noNewLine
			Write-host "$long "-ForegroundColor Green -noNewLine
			Write-host " cadencia: " -noNewLine
			Write-host "$cadf" -ForegroundColor Green
			
		}
	}
}
Write-Host "`nSalvando novo arquivo " -noNewLine
Write-Host "$currentDir\$newgpx" -ForegroundColor Yellow
if(-not(Test-Path $dirDest)){ 
	New-Item -ItemType directory -Path $dirDest > $null
}

$gpx.Save("$currentDir\$newgpx")
Write-Host "Movendo arquivos $csvFile e $gpxFile para o diretorio $dirDest`n"
$csvFile |Move-Item -Destination $dirDest -Force
$gpxFile |Move-Item -Destination $dirDest -Force
#get-content "C:\Users\Lima\Desktop\GPX\test.xml"

#Line 200