
Function DegToRad(deg) As Double
 Pi = 3.14159265358979
 DegToRad = (deg / 180 * Pi)
End Function

Function RadToDeg(rad) As Double
 Pi = 3.14159265358979
 RadToDeg = (rad / Pi * 180)
End Function


Function UTMCentralMeridian(zone) As Double
  cmeridian = DegToRad(-183 + (zone * 6))
  UTMCentralMeridian = cmeridian
End Function

Function FootpointLatitude(y) As Double
  sm_a = 6378137
  sm_b = 6356752.314
  
  n = (sm_a - sm_b) / (sm_a + sm_b)
  
  alpha_ = ((sm_a + sm_b) / 2) * (1 + ((n ^ 2) / 4) + ((n ^ 4) / 64))
  
  y_ = y / alpha_
  
  beta_ = (3 * n / 2) + (-27 * (n ^ 3) / 32) + (269 * (n ^ 5) / 512)
  
  gamma_ = (21 * (n ^ 2) / 16) + (-55 * (n ^ 4) / 32)
  
  delta_ = (151 * (n ^ 3) / 96) + (-417 * (n ^ 5) / 128)
  
  epsilon_ = (1097 * (n ^ 4) / 512)
  
  result = y_ + (beta_ * Sin(2 * y_)) + (gamma_ * Sin(4 * y_)) + (delta_ * Sin(6 * y_)) + (epsilon_ * Sin(8 * y_))
  
  FootpointLatitude = result
End Function


Function MapXYToLatLon(x, y, lambda0) As Collection
  Dim coordinates As Collection
  Set coordinates = New Collection
  
  sm_a = 6378137
  sm_b = 6356752.314
  
  phif = FootpointLatitude(y)
  
  ep2 = ((sm_a ^ 2) - (sm_b ^ 2)) / (sm_b ^ 2)
  
  cf = Cos(phif)
  
  nuf2 = ep2 * (cf ^ 2)
  
  Nf = (sm_a ^ 2) / (sm_b * ((1 + nuf2) ^ (1 / 2)))
  Nfpow = Nf
  
  tf = Tan(phif)
  tf2 = tf * tf
  tf4 = tf2 * tf2
  
  x1frac = 1 / (Nfpow * cf)
  
  Nfpow = Nfpow * Nf           ' now equals Nf**2
  x2frac = tf / (2 * Nfpow)
  
  Nfpow = Nfpow * Nf           ' now equals Nf**3
  x3frac = 1 / (6 * Nfpow * cf)
  
  Nfpow = Nfpow * Nf           ' now equals Nf**4
  x4frac = tf / (24 * Nfpow)
  
  Nfpow = Nfpow * Nf           ' now equals Nf**5
  x5frac = 1 / (120 * Nfpow * cf)
  
  Nfpow = Nfpow * Nf           ' now equals Nf**6
  x6frac = tf / (720 * Nfpow)
  
  Nfpow = Nfpow * Nf           ' now equals Nf**7
  x7frac = 1 / (5040 * Nfpow * cf)
  
  Nfpow = Nfpow * Nf           ' now equals Nf**8
  x8frac = tf / (40320 * Nfpow)
  
  x2poly = -1 - nuf2
  x3poly = -1 - 2 * tf2 - nuf2
  
  x4poly = 5 + 3 * tf2 + 6 * nuf2 - 6 * tf2 * nuf2 - 3 * (nuf2 * nuf2) - 9 * tf2 * (nuf2 * nuf2)
  x5poly = 5 + 28 * tf2 + 24 * tf4 + 6 * nuf2 + 8 * tf2 * nuf2
  x6poly = -61 - 90 * tf2 - 45 * tf4 - 107 * nuf2 + 162 * tf2 * nuf2
  x7poly = -61 - 662 * tf2 - 1320 * tf4 - 720 * (tf4 * tf2)
  x8poly = 1385 + 3633 * tf2 + 4095 * tf4 + 1575 * (tf4 * tf2)
  
  ' Calculate latitude
  latitude = phif + x2frac * x2poly * (x * x) + x4frac * x4poly * (x ^ 4) + x6frac * x6poly * (x ^ 6) + x8frac * x8poly * (x ^ 8)
    
  ' Calculate longitude
  longitude = lambda0 + x1frac * x + x3frac * x3poly * (x ^ 3) + x5frac * x5poly * (x ^ 5) + x7frac * x7poly * (x ^ 7)
    
  coordinates.Add RadToDeg(latitude)
  coordinates.Add RadToDeg(longitude)
  
  Debug.Print coordinates(1)
  Debug.Print coordinates(2)

  Set MapXYToLatLon = coordinates
End Function


Function UTMXYToLatLon(ByVal easting As Double, ByVal northing As Double, ByVal utm_zone As Integer, ByVal is_north As Boolean) As Collection
  UTMScaleFactor = 0.9996
  
  easting = easting - 500000
  easting = easting / UTMScaleFactor
    
  If is_north = False Then
    northing = northing - 10000000
  End If
    
      
  northing = northing / UTMScaleFactor
  
  cmeridian = UTMCentralMeridian(utm_zone)
  
  Set UTMXYToLatLon = MapXYToLatLon(easting, northing, cmeridian)
End Function
