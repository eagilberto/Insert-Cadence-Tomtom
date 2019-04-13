# Insert-Cadence-Tomtom
Script powershell to insert step cadence on activities to strava

1) Export activity from https://mysports.tomtom.com formats .csv and .gpx to same directory of script
2) Execute .\InsertCadence.ps1 and select your activity
3) The element <gpxtpx:cad>cadence</gpxtpx:cad> will be added from .csv (cycles column) to .gpx

      <trkpt lat="-16.657631" lon="-49.376084">
        <ele>828.7</ele>
        <time>2018-09-09T21:14:49.000Z</time>
        <extensions>
          <gpxtpx:TrackPointExtension>
            <gpxtpx:hr>176</gpxtpx:hr>
            <gpxtpx:cad>84</gpxtpx:cad>
          </gpxtpx:TrackPointExtension>
        </extensions>
      </trkpt>


3) Import the new file New_run-YYYYMMDDThhmmss.gpx to Strava 
 
4) My Marathon with cadence: https://www.strava.com/activities/2285616271
  
 
