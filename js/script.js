// Embedded Calendar Dataset from Excel
    const CALENDAR_DATA = [{"week_num": 1, "excel_row": 3, "start_date": "2027-01-01", "end_date": "2027-01-03", "date_formatted": "Jan 01 – 03, 2027", "month": "January", "month_idx": 1, "quarter": "Q1", "status": "New Year", "status_desc": "Pre-Season Window", "event_count": 0, "events": []}, {"week_num": 2, "excel_row": 4, "start_date": "2027-01-08", "end_date": "2027-01-10", "date_formatted": "Jan 08 – 10, 2027", "month": "January", "month_idx": 1, "quarter": "Q1", "status": "Race Weekend", "status_desc": "🏁 1 Championship Event", "event_count": 1, "events": [{"col": 13, "row": 4, "raw_text": "Hampton Downs", "series": "FROT", "sub_series": "Formula Regional Oceania", "badge": "FROT (NZ)", "color_cat": "frot", "excel_note": "Theme 8 (Purple): Formula Regional Oceania Trophy", "title": "Castrol Toyota FR Oceania - Hampton Downs", "circuit": "Hampton Downs Motorsport Park", "city": "Waikato", "country": "New Zealand", "flag": "🇳🇿", "notes": "", "theme": 8, "rgb": null}]}, {"week_num": 3, "excel_row": 5, "start_date": "2027-01-15", "end_date": "2027-01-17", "date_formatted": "Jan 15 – 17, 2027", "month": "January", "month_idx": 1, "quarter": "Q1", "status": "Multi-Series Weekend", "status_desc": "⚡ 3 Championships in Action", "event_count": 3, "events": [{"col": 4, "row": 5, "raw_text": "Mexico City", "series": "Formula E", "sub_series": "FE World Championship", "badge": "E-Prix", "color_cat": "fe", "excel_note": "Single E-Prix Race", "title": "Mexico City E-Prix", "circuit": "Autódromo Hermanos Rodríguez", "city": "Mexico City", "country": "Mexico", "flag": "🇲🇽", "notes": "", "theme": null, "rgb": "00000000"}, {"col": 5, "row": 5, "raw_text": "Jerez", "series": "Asian Le Mans", "sub_series": "Asian Le Mans Series", "badge": "ALMS", "color_cat": "alms", "excel_note": "Light Yellow: Asian Le Mans Winter Round", "title": "Asian Le Mans 4 Hours of Jerez", "circuit": "Circuito de Jerez-Ángel Nieto", "city": "Jerez de la Frontera", "country": "Spain", "flag": "🇪🇸", "notes": "", "theme": null, "rgb": "FFFFFFCC"}, {"col": 13, "row": 5, "raw_text": "Taupo", "series": "FROT", "sub_series": "Formula Regional Oceania", "badge": "FROT (NZ)", "color_cat": "frot", "excel_note": "Theme 8 (Purple): Formula Regional Oceania Trophy", "title": "Castrol Toyota FR Oceania - Taupo", "circuit": "Taupo International Motorsport Park", "city": "Taupo", "country": "New Zealand", "flag": "🇳🇿", "notes": "", "theme": 8, "rgb": null}]}, {"week_num": 4, "excel_row": 6, "start_date": "2027-01-22", "end_date": "2027-01-24", "date_formatted": "Jan 22 – 24, 2027", "month": "January", "month_idx": 1, "quarter": "Q1", "status": "Multi-Series Weekend", "status_desc": "⚡ 3 Championships in Action", "event_count": 3, "events": [{"col": 12, "row": 6, "raw_text": "Monaco", "series": "WRC", "sub_series": "World Rally Championship", "badge": "WRC", "color_cat": "wrc", "excel_note": "FIA World Rally Championship", "title": "WRC 95th Rallye Automobile Monte-Carlo", "circuit": "French Alps & Col de Turini", "city": "Gap / Monaco", "country": "Monaco", "flag": "🇲🇨", "notes": "Tarmac, Ice & Snow • Season Opener", "theme": null, "rgb": "00000000"}, {"col": 13, "row": 6, "raw_text": "Teretonga Park", "series": "FROT", "sub_series": "Formula Regional Oceania", "badge": "FROT (NZ)", "color_cat": "frot", "excel_note": "Theme 8 (Purple): Formula Regional Oceania Trophy", "title": "Castrol Toyota FR Oceania - Teretonga Park", "circuit": "Teretonga Park", "city": "Invercargill", "country": "New Zealand", "flag": "🇳🇿", "notes": "World's Southernmost Permanent Race Track", "theme": 8, "rgb": null}, {"col": 14, "row": 6, "raw_text": "Estoril", "series": "FWS", "sub_series": "Formula Winter Series", "badge": "FWS (Winter)", "color_cat": "fws", "excel_note": "Light Yellow: Formula Winter Series (Iberia)", "title": "FWS Estoril", "circuit": "Estoril", "city": "", "country": "", "flag": "🏁", "notes": "", "theme": null, "rgb": "FFFFFFCC"}]}, {"week_num": 5, "excel_row": 7, "start_date": "2027-01-29", "end_date": "2027-01-31", "date_formatted": "Jan 29 – 31, 2027", "month": "January", "month_idx": 1, "quarter": "Q1", "status": "Race Weekend", "status_desc": "🏁 2 Championship Events", "event_count": 2, "events": [{"col": 10, "row": 7, "raw_text": "Daytona", "series": "IMSA", "sub_series": "IMSA SportsCar Championship", "badge": "IMSA", "color_cat": "imsa", "excel_note": "IMSA WeatherTech SportsCar Championship", "title": "65th Rolex 24 at Daytona", "circuit": "Daytona International Speedway", "city": "Daytona Beach, FL", "country": "United States", "flag": "🇺🇸", "notes": "North American 24h Classic • IMSA Opener", "theme": null, "rgb": "00000000"}, {"col": 13, "row": 7, "raw_text": "Highlands", "series": "FROT", "sub_series": "Formula Regional Oceania", "badge": "FROT (NZ)", "color_cat": "frot", "excel_note": "Theme 8 (Purple): Formula Regional Oceania Trophy", "title": "Castrol Toyota FR Oceania - Highlands (NZ GP)", "circuit": "Highlands Motorsport Park", "city": "Cromwell, Otago", "country": "New Zealand", "flag": "🇳🇿", "notes": "71st New Zealand Grand Prix", "theme": 8, "rgb": null}]}, {"week_num": 6, "excel_row": 8, "start_date": "2027-02-05", "end_date": "2027-02-07", "date_formatted": "Feb 05 – 07, 2027", "month": "February", "month_idx": 2, "quarter": "Q1", "status": "Race Weekend", "status_desc": "🏁 2 Championship Events", "event_count": 2, "events": [{"col": 4, "row": 8, "raw_text": "Austin", "series": "Formula E", "sub_series": "FE World Championship", "badge": "E-Prix", "color_cat": "fe", "excel_note": "Single E-Prix Race", "title": "Austin E-Prix", "circuit": "Circuit of the Americas (COTA)", "city": "Austin, TX", "country": "United States", "flag": "🇺🇸", "notes": "", "theme": null, "rgb": "00000000"}, {"col": 13, "row": 8, "raw_text": "Portimão", "series": "FWS", "sub_series": "Formula Winter Series", "badge": "FWS (Winter)", "color_cat": "fws", "excel_note": "Light Yellow: Formula Winter Series (Iberia)", "title": "Formula Winter Series - Portimão", "circuit": "Autódromo Internacional do Algarve", "city": "Portimão", "country": "Portugal", "flag": "🇵🇹", "notes": "", "theme": null, "rgb": "FFFFFFCC"}]}, {"week_num": 7, "excel_row": 9, "start_date": "2027-02-12", "end_date": "2027-02-14", "date_formatted": "Feb 12 – 14, 2027", "month": "February", "month_idx": 2, "quarter": "Q1", "status": "Race Weekend", "status_desc": "🏁 2 Championship Events", "event_count": 2, "events": [{"col": 6, "row": 9, "raw_text": "Bathurst 12 Hour", "series": "IGTC", "sub_series": "Intercontinental GT Challenge", "badge": "IGTC", "color_cat": "igtc", "excel_note": "Theme 3 (Grey): Intercontinental GT Classic", "title": "Repco Bathurst 12 Hour", "circuit": "Mount Panorama Circuit", "city": "Bathurst, NSW", "country": "Australia", "flag": "🇦🇺", "notes": "Mount Panorama Classic • IGTC Opener", "theme": 3, "rgb": null}, {"col": 13, "row": 9, "raw_text": "Jerez", "series": "FWS", "sub_series": "Formula Winter Series", "badge": "FWS (Winter)", "color_cat": "fws", "excel_note": "Light Yellow: Formula Winter Series (Iberia)", "title": "Formula Winter Series - Jerez", "circuit": "Circuito de Jerez-Ángel Nieto", "city": "Jerez de la Frontera", "country": "Spain", "flag": "🇪🇸", "notes": "", "theme": null, "rgb": "FFFFFFCC"}]}, {"week_num": 8, "excel_row": 10, "start_date": "2027-02-19", "end_date": "2027-02-21", "date_formatted": "Feb 19 – 21, 2027", "month": "February", "month_idx": 2, "quarter": "Q1", "status": "Multi-Series Weekend", "status_desc": "⚡ 3 Championships in Action", "event_count": 3, "events": [{"col": 4, "row": 10, "raw_text": "Miami", "series": "Formula E", "sub_series": "FE World Championship", "badge": "E-Prix", "color_cat": "fe", "excel_note": "Single E-Prix Race", "title": "Miami E-Prix", "circuit": "Homestead-Miami / Biscayne Bay", "city": "Miami, FL", "country": "United States", "flag": "🇺🇸", "notes": "", "theme": null, "rgb": "00000000"}, {"col": 5, "row": 10, "raw_text": "Portimao", "series": "Asian Le Mans", "sub_series": "Asian Le Mans Series", "badge": "ALMS", "color_cat": "alms", "excel_note": "Light Yellow: Asian Le Mans Winter Round", "title": "Asian Le Mans 4 Hours of Portimão", "circuit": "Autódromo Internacional do Algarve", "city": "Portimão", "country": "Portugal", "flag": "🇵🇹", "notes": "", "theme": null, "rgb": "FFFFFFCC"}, {"col": 12, "row": 10, "raw_text": "Sweden", "series": "WRC", "sub_series": "World Rally Championship", "badge": "WRC", "color_cat": "wrc", "excel_note": "FIA World Rally Championship", "title": "WRC Rally Sweden", "circuit": "Snow & Ice Stages of Västerbotten", "city": "Umeå", "country": "Sweden", "flag": "🇸🇪", "notes": "Pure Snow & Studded Tires (-20°C)", "theme": null, "rgb": "00000000"}]}, {"week_num": 9, "excel_row": 11, "start_date": "2027-02-26", "end_date": "2027-02-28", "date_formatted": "Feb 26 – 28, 2027", "month": "February", "month_idx": 2, "quarter": "Q1", "status": "Race Weekend", "status_desc": "🏁 1 Championship Event", "event_count": 1, "events": [{"col": 13, "row": 11, "raw_text": "Barcelona", "series": "FWS", "sub_series": "Formula Winter Series", "badge": "FWS (Winter)", "color_cat": "fws", "excel_note": "Light Yellow: Formula Winter Series (Iberia)", "title": "Formula Winter Series - Barcelona", "circuit": "Circuit de Barcelona-Catalunya", "city": "Montmeló", "country": "Spain", "flag": "🇪🇸", "notes": "", "theme": null, "rgb": "FFFFFFCC"}]}, {"week_num": 10, "excel_row": 12, "start_date": "2027-03-05", "end_date": "2027-03-07", "date_formatted": "Mar 05 – 07, 2027", "month": "March", "month_idx": 3, "quarter": "Q1", "status": "Race Weekend", "status_desc": "🏁 1 Championship Event", "event_count": 1, "events": [{"col": 13, "row": 12, "raw_text": "Valencia", "series": "FWS", "sub_series": "Formula Winter Series", "badge": "FWS (Winter)", "color_cat": "fws", "excel_note": "Light Yellow: Formula Winter Series (Iberia)", "title": "Formula Winter Series - Valencia Finale", "circuit": "Circuit Ricardo Tormo", "city": "Cheste, Valencia", "country": "Spain", "flag": "🇪🇸", "notes": "FWS Championship Finale", "theme": null, "rgb": "FFFFFFCC"}]}, {"week_num": 11, "excel_row": 13, "start_date": "2027-03-12", "end_date": "2027-03-14", "date_formatted": "Mar 12 – 14, 2027", "month": "March", "month_idx": 3, "quarter": "Q1", "status": "Multi-Series Weekend", "status_desc": "⚡ 3 Championships in Action", "event_count": 3, "events": [{"col": 3, "row": 13, "raw_text": "Bahrain (s)", "series": "Formula 1", "sub_series": "F1 • F2 & F3 Support", "badge": "F1 + F2 + F3 • Sprint", "color_cat": "f1-f3", "excel_note": "Green Cell: F3 & F2 Support Races Included • Sprint Weekend (s)", "title": "Formula 1 Bahrain Grand Prix", "circuit": "Bahrain International Circuit", "city": "Sakhir", "country": "Bahrain", "flag": "🇧🇭", "notes": "F1 Sprint Weekend • Desert Opener", "theme": 6, "rgb": null}, {"col": 4, "row": 13, "raw_text": "Sao Paulo", "series": "Formula E", "sub_series": "FE World Championship", "badge": "E-Prix", "color_cat": "fe", "excel_note": "Single E-Prix Race", "title": "São Paulo E-Prix", "circuit": "Sambadrome Anhembi", "city": "São Paulo", "country": "Brazil", "flag": "🇧🇷", "notes": "", "theme": null, "rgb": "00000000"}, {"col": 5, "row": 13, "raw_text": "Barcelona", "series": "ELMS", "sub_series": "European Le Mans Series", "badge": "ELMS", "color_cat": "elms", "excel_note": "Theme 5 (Orange): European Le Mans Series", "title": "ELMS 4 Hours of Barcelona", "circuit": "Circuit de Barcelona-Catalunya", "city": "Montmeló", "country": "Spain", "flag": "🇪🇸", "notes": "", "theme": 5, "rgb": null}]}, {"week_num": 12, "excel_row": 14, "start_date": "2027-03-19", "end_date": "2027-03-21", "date_formatted": "Mar 19 – 21, 2027", "month": "March", "month_idx": 3, "quarter": "Q1", "status": "Multi-Series Weekend", "status_desc": "⚡ 3 Championships in Action", "event_count": 3, "events": [{"col": 3, "row": 14, "raw_text": "Saudi Arabia", "series": "Formula 1", "sub_series": "F1 • F2 & F3 Support", "badge": "F1 + F2 + F3", "color_cat": "f1-f3", "excel_note": "Green Cell: F3 & F2 Support Races Included", "title": "Formula 1 Saudi Arabian Grand Prix", "circuit": "Jeddah Corniche Circuit", "city": "Jeddah", "country": "Saudi Arabia", "flag": "🇸🇦", "notes": "Fastest Street Circuit • Night Race", "theme": 6, "rgb": null}, {"col": 10, "row": 14, "raw_text": "Sebring", "series": "IMSA", "sub_series": "IMSA SportsCar Championship", "badge": "IMSA", "color_cat": "imsa", "excel_note": "IMSA WeatherTech SportsCar Championship", "title": "75th Mobil 1 Twelve Hours of Sebring", "circuit": "Sebring International Raceway", "city": "Sebring, FL", "country": "United States", "flag": "🇺🇸", "notes": "Historic Bumpy 12h Classic • Florida 36h", "theme": null, "rgb": "00000000"}, {"col": 12, "row": 14, "raw_text": "Kenya", "series": "WRC", "sub_series": "World Rally Championship", "badge": "WRC", "color_cat": "wrc", "excel_note": "FIA World Rally Championship", "title": "WRC Safari Rally Kenya", "circuit": "Great Rift Valley Wilderness Stages", "city": "Naivasha", "country": "Kenya", "flag": "🇰🇪", "notes": "Brutal Rough Gravel, Fesh-Fesh & Wildlife", "theme": null, "rgb": "00000000"}]}, {"week_num": 13, "excel_row": 15, "start_date": "2027-03-26", "end_date": "2027-03-28", "date_formatted": "Mar 26 – 28, 2027", "month": "March", "month_idx": 3, "quarter": "Q1", "status": "Race Weekend", "status_desc": "🏁 1 Championship Event", "event_count": 1, "events": [{"col": 5, "row": 15, "raw_text": "Qatar 1812km", "series": "FIA WEC", "sub_series": "World Endurance Championship", "badge": "WEC", "color_cat": "wec", "excel_note": "Theme 3 (Grey): FIA World Endurance Championship", "title": "FIA WEC Qatar 1812km", "circuit": "Lusail International Circuit", "city": "Lusail", "country": "Qatar", "flag": "🇶🇦", "notes": "WEC Season Opener • 10 Hours / 1812km", "theme": 3, "rgb": null}]}, {"week_num": 14, "excel_row": 16, "start_date": "2027-04-02", "end_date": "2027-04-04", "date_formatted": "Apr 02 – 04, 2027", "month": "April", "month_idx": 4, "quarter": "Q2", "status": "Race Weekend", "status_desc": "🏁 2 Championship Events", "event_count": 2, "events": [{"col": 3, "row": 16, "raw_text": "Austraila (s)", "series": "Formula 1", "sub_series": "F1 • F2 & F3 Support", "badge": "F1 + F2 + F3 • Sprint", "color_cat": "f1-f3", "excel_note": "Green Cell: F3 & F2 Support Races Included • Sprint Weekend (s)", "title": "Formula 1 Australian Grand Prix", "circuit": "Albert Park Circuit", "city": "Melbourne", "country": "Australia", "flag": "🇦🇺", "notes": "F1 Sprint Weekend • Lakeside Circuit", "theme": 6, "rgb": null}, {"col": 6, "row": 16, "raw_text": "Sepang", "series": "GTWC Asia", "sub_series": "GT World Challenge Asia", "badge": "GTWC Asia", "color_cat": "gtwca", "excel_note": "Theme 9 (Green): GT World Challenge Asia", "title": "GTWC Asia Sepang", "circuit": "Sepang International Circuit", "city": "Sepang, Selangor", "country": "Malaysia", "flag": "🇲🇾", "notes": "", "theme": 9, "rgb": null}]}, {"week_num": 15, "excel_row": 17, "start_date": "2027-04-09", "end_date": "2027-04-11", "date_formatted": "Apr 09 – 11, 2027", "month": "April", "month_idx": 4, "quarter": "Q2", "status": "Mega Super Weekend", "status_desc": "🔥 5 Major Championships Racing!", "event_count": 5, "events": [{"col": 3, "row": 17, "raw_text": "Japan (s)", "series": "Formula 1", "sub_series": "F1 Standalone", "badge": "F1 Exclusive • Sprint", "color_cat": "f1-only", "excel_note": "Uncolored: F1 Headline Only • Sprint Weekend (s)", "title": "Formula 1 Japanese Grand Prix", "circuit": "Suzuka International Racing Course", "city": "Suzuka", "country": "Japan", "flag": "🇯🇵", "notes": "F1 Sprint Weekend • Figure-Eight Classic", "theme": null, "rgb": "00000000"}, {"col": 5, "row": 17, "raw_text": "6 Hours of IMOLA", "series": "FIA WEC", "sub_series": "World Endurance Championship", "badge": "WEC", "color_cat": "wec", "excel_note": "Theme 3 (Grey): FIA World Endurance Championship", "title": "FIA WEC 6 Hours of Imola", "circuit": "Autodromo Enzo e Dino Ferrari", "city": "Imola, Emilia-Romagna", "country": "Italy", "flag": "🇮🇹", "notes": "", "theme": 3, "rgb": null}, {"col": 12, "row": 17, "raw_text": "Croatia", "series": "WRC", "sub_series": "World Rally Championship", "badge": "WRC", "color_cat": "wrc", "excel_note": "FIA World Rally Championship", "title": "WRC Croatia Rally", "circuit": "Asphalt Stages of Karlovac & Zagreb", "city": "Zagreb", "country": "Croatia", "flag": "🇭🇷", "notes": "Technical Tarmac with Grip Surface Changes", "theme": null, "rgb": "00000000"}, {"col": 13, "row": 17, "raw_text": "Portugal", "series": "FRECA", "sub_series": "Formula Regional European Championship", "badge": "FRECA", "color_cat": "freca", "excel_note": "Theme 5 (Orange): Formula Regional Europe (Alpine)", "title": "FRECA Portugal (Portimão)", "circuit": "Autódromo Internacional do Algarve", "city": "Portimão", "country": "Portugal", "flag": "🇵🇹", "notes": "", "theme": 5, "rgb": null}, {"col": 15, "row": 17, "raw_text": "Jarama", "series": "Spanish F4", "sub_series": "F4 Spanish Championship", "badge": "Spanish F4", "color_cat": "f4-es", "excel_note": "Light Yellow: Campeonato de España de F4", "title": "Spanish F4 - Jarama", "circuit": "Circuito del Jarama", "city": "San Sebastián de los Reyes", "country": "Spain", "flag": "🇪🇸", "notes": "", "theme": null, "rgb": "FFFFFFCC"}]}, {"week_num": 16, "excel_row": 18, "start_date": "2027-04-16", "end_date": "2027-04-18", "date_formatted": "Apr 16 – 18, 2027", "month": "April", "month_idx": 4, "quarter": "Q2", "status": "Multi-Series Weekend", "status_desc": "⚡ 4 Championships in Action", "event_count": 4, "events": [{"col": 3, "row": 18, "raw_text": "China", "series": "Formula 1", "sub_series": "F1 Standalone", "badge": "F1 Exclusive", "color_cat": "f1-only", "excel_note": "Uncolored: F1 Headline Only", "title": "Formula 1 Chinese Grand Prix", "circuit": "Shanghai International Circuit", "city": "Shanghai", "country": "China", "flag": "🇨🇳", "notes": "Iconic 1.2km Back Straight", "theme": null, "rgb": "00000000"}, {"col": 4, "row": 18, "raw_text": "Sanya", "series": "Formula E", "sub_series": "FE World Championship", "badge": "E-Prix", "color_cat": "fe", "excel_note": "Single E-Prix Race", "title": "Sanya E-Prix", "circuit": "Haitang Bay Street Circuit", "city": "Sanya, Hainan", "country": "China", "flag": "🇨🇳", "notes": "", "theme": null, "rgb": "00000000"}, {"col": 6, "row": 18, "raw_text": "Paul Ricard 6 Hour", "series": "GTWCE", "sub_series": "GT World Challenge Europe", "badge": "GTWCE Enduro", "color_cat": "gtwce", "excel_note": "Theme 5 (Orange): GTWCE Endurance Cup", "title": "GTWCE Paul Ricard 1000km (6 Hours)", "circuit": "Circuit Paul Ricard", "city": "Le Castellet", "country": "France", "flag": "🇫🇷", "notes": "GTWCE Endurance Cup Opener", "theme": 5, "rgb": null}, {"col": 10, "row": 18, "raw_text": "Long Beach", "series": "IMSA", "sub_series": "IMSA SportsCar Championship", "badge": "IMSA", "color_cat": "imsa", "excel_note": "IMSA WeatherTech SportsCar Championship", "title": "Acura Grand Prix of Long Beach", "circuit": "Streets of Long Beach", "city": "Long Beach, CA", "country": "United States", "flag": "🇺🇸", "notes": "100-Minute Street Sprint", "theme": null, "rgb": "00000000"}]}, {"week_num": 17, "excel_row": 19, "start_date": "2027-04-23", "end_date": "2027-04-25", "date_formatted": "Apr 23 – 25, 2027", "month": "April", "month_idx": 4, "quarter": "Q2", "status": "Multi-Series Weekend", "status_desc": "⚡ 3 Championships in Action", "event_count": 3, "events": [{"col": 5, "row": 19, "raw_text": "6 Hours of Silverstone", "series": "FIA WEC", "sub_series": "World Endurance Championship", "badge": "WEC", "color_cat": "wec", "excel_note": "Theme 3 (Grey): FIA World Endurance Championship", "title": "FIA WEC 6 Hours of Silverstone", "circuit": "Silverstone Circuit", "city": "Silverstone", "country": "United Kingdom", "flag": "🇬🇧", "notes": "", "theme": 3, "rgb": null}, {"col": 11, "row": 19, "raw_text": "Red Bull Ring", "series": "DTM", "sub_series": "Deutsche Tourenwagen Masters", "badge": "DTM", "color_cat": "dtm", "excel_note": "Deutsche Tourenwagen Masters (2 Races)", "title": "DTM Red Bull Ring", "circuit": "Red Bull Ring", "city": "", "country": "", "flag": "🏁", "notes": "", "theme": null, "rgb": "00000000"}, {"col": 13, "row": 19, "raw_text": "Austria", "series": "FRECA", "sub_series": "Formula Regional European Championship", "badge": "FRECA", "color_cat": "freca", "excel_note": "Theme 5 (Orange): Formula Regional Europe (Alpine)", "title": "FRECA Austria (Red Bull Ring)", "circuit": "Red Bull Ring", "city": "Spielberg", "country": "Austria", "flag": "🇦🇹", "notes": "", "theme": 5, "rgb": null}]}, {"week_num": 18, "excel_row": 20, "start_date": "2027-04-30", "end_date": "2027-05-02", "date_formatted": "Apr 30 – May 02, 2027", "month": "April", "month_idx": 4, "quarter": "Q2", "status": "Mega Super Weekend", "status_desc": "🔥 6 Major Championships Racing!", "event_count": 6, "events": [{"col": 3, "row": 20, "raw_text": "Miami", "series": "Formula 1", "sub_series": "F1 Standalone", "badge": "F1 Exclusive", "color_cat": "f1-only", "excel_note": "Uncolored: F1 Headline Only", "title": "Formula 1 Crypto.com Miami Grand Prix", "circuit": "Miami International Autodrome", "city": "Miami Gardens, FL", "country": "United States", "flag": "🇺🇸", "notes": "Hard Rock Stadium Campus", "theme": null, "rgb": "00000000"}, {"col": 4, "row": 20, "raw_text": "Monaco*2", "series": "Formula E", "sub_series": "FE World Championship", "badge": "Double Header (*2)", "color_cat": "fe", "excel_note": "Double Header: 2 E-Prix Races", "title": "Monaco E-Prix (Double Header)", "circuit": "Circuit de Monaco", "city": "Monte Carlo", "country": "Monaco", "flag": "🇲🇨", "notes": "Rounds 1 & 2 • Port Hercule", "theme": null, "rgb": "00000000"}, {"col": 5, "row": 20, "raw_text": "Le Castellet", "series": "ELMS", "sub_series": "European Le Mans Series", "badge": "ELMS", "color_cat": "elms", "excel_note": "Theme 5 (Orange): European Le Mans Series", "title": "ELMS 4 Hours of Le Castellet", "circuit": "Circuit Paul Ricard", "city": "Le Castellet, Var", "country": "France", "flag": "🇫🇷", "notes": "", "theme": 5, "rgb": null}, {"col": 6, "row": 20, "raw_text": "Brands Hatch (s)", "series": "GTWCE", "sub_series": "GT World Challenge Europe", "badge": "GTWCE Sprint (s)", "color_cat": "gtwce", "excel_note": "Theme 5 (Orange): GTWCE Sprint Cup (s)", "title": "GTWCE Brands Hatch Sprint Cup", "circuit": "Brands Hatch Circuit", "city": "Fawkham, Kent", "country": "United Kingdom", "flag": "🇬🇧", "notes": "GTWCE Sprint Cup (2x 60min Races)", "theme": 5, "rgb": null}, {"col": 8, "row": 20, "raw_text": "Thailand", "series": "GTWC Asia", "sub_series": "GT World Challenge Asia", "badge": "GTWC Asia", "color_cat": "gtwca", "excel_note": "Theme 9 (Green): GT World Challenge Asia", "title": "GTWC Asia Chang (Buriram)", "circuit": "Chang International Circuit", "city": "Buriram", "country": "Thailand", "flag": "🇹🇭", "notes": "", "theme": 9, "rgb": null}, {"col": 10, "row": 20, "raw_text": "Laguna Seca", "series": "IMSA", "sub_series": "IMSA SportsCar Championship", "badge": "IMSA", "color_cat": "imsa", "excel_note": "IMSA WeatherTech SportsCar Championship", "title": "Motul Course de Monterey at Laguna Seca", "circuit": "WeatherTech Raceway Laguna Seca", "city": "Salinas / Monterey, CA", "country": "United States", "flag": "🇺🇸", "notes": "Iconic Corkscrew • 2h 40m", "theme": null, "rgb": "00000000"}]}, {"week_num": 19, "excel_row": 21, "start_date": "2027-05-07", "end_date": "2027-05-09", "date_formatted": "May 07 – 09, 2027", "month": "May", "month_idx": 5, "quarter": "Q2", "status": "Mega Super Weekend", "status_desc": "🔥 6 Major Championships Racing!", "event_count": 6, "events": [{"col": 4, "row": 21, "raw_text": "Berlin*2", "series": "Formula E", "sub_series": "FE World Championship", "badge": "Double Header (*2)", "color_cat": "fe", "excel_note": "Double Header: 2 E-Prix Races", "title": "Berlin E-Prix (Double Header)", "circuit": "Tempelhof Airport Street Circuit", "city": "Berlin", "country": "Germany", "flag": "🇩🇪", "notes": "Rounds 1 & 2 • Apron Concrete", "theme": null, "rgb": "00000000"}, {"col": 6, "row": 21, "raw_text": "Texas 8 Hour", "series": "IGTC", "sub_series": "Intercontinental GT Challenge", "badge": "IGTC", "color_cat": "igtc", "excel_note": "Theme 3 (Grey): Intercontinental GT Classic", "title": "IGTC Texas 8 Hour at COTA", "circuit": "Circuit of the Americas (COTA)", "city": "Austin, TX", "country": "United States", "flag": "🇺🇸", "notes": "IGTC American Round", "theme": 3, "rgb": null}, {"col": 11, "row": 21, "raw_text": "Oschersleben", "series": "DTM", "sub_series": "Deutsche Tourenwagen Masters", "badge": "DTM", "color_cat": "dtm", "excel_note": "Deutsche Tourenwagen Masters (2 Races)", "title": "DTM Motorsport Arena Oschersleben", "circuit": "Motorsport Arena Oschersleben", "city": "Oschersleben", "country": "Germany", "flag": "🇩🇪", "notes": "Races 1 & 2", "theme": null, "rgb": "00000000"}, {"col": 12, "row": 21, "raw_text": "Spain", "series": "WRC", "sub_series": "World Rally Championship", "badge": "WRC", "color_cat": "wrc", "excel_note": "FIA World Rally Championship", "title": "WRC Rally Islas Canarias (Spain)", "circuit": "Volcanic Mountain Asphalt", "city": "Las Palmas, Gran Canaria", "country": "Spain", "flag": "🇪🇸", "notes": "Winding Mountain Tarmac", "theme": null, "rgb": "00000000"}, {"col": 15, "row": 21, "raw_text": "Oschersleben", "series": "German F4", "sub_series": "German F4 / ADAC", "badge": "German F4", "color_cat": "f4-de", "excel_note": "Theme 3 (Grey): ADAC Formula 4 (DTM Support)", "title": "German F4 Oschersleben", "circuit": "Oschersleben", "city": "", "country": "", "flag": "🏁", "notes": "", "theme": 3, "rgb": null}, {"col": 17, "row": 21, "raw_text": "Misano", "series": "Italian F4", "sub_series": "Italian F4 / Euro 4", "badge": "Italian F4", "color_cat": "f4-it", "excel_note": "Theme 6 (Green): Italian F4 / Euro 4 Championship", "title": "Italian F4 - Misano", "circuit": "Misano World Circuit Marco Simoncelli", "city": "Misano Adriatico", "country": "Italy", "flag": "🇮🇹", "notes": "", "theme": 6, "rgb": null}]}, {"week_num": 20, "excel_row": 22, "start_date": "2027-05-14", "end_date": "2027-05-16", "date_formatted": "May 14 – 16, 2027", "month": "May", "month_idx": 5, "quarter": "Q2", "status": "Race Weekend", "status_desc": "🏁 2 Championship Events", "event_count": 2, "events": [{"col": 5, "row": 22, "raw_text": "6 Hours of SPA", "series": "FIA WEC", "sub_series": "World Endurance Championship", "badge": "WEC", "color_cat": "wec", "excel_note": "Theme 3 (Grey): FIA World Endurance Championship", "title": "FIA WEC 6 Hours of Spa-Francorchamps", "circuit": "Circuit de Spa-Francorchamps", "city": "Stavelot", "country": "Belgium", "flag": "🇧🇪", "notes": "Ardennes Dress Rehearsal for Le Mans", "theme": 3, "rgb": null}, {"col": 13, "row": 22, "raw_text": "Silverstone", "series": "FRECA", "sub_series": "Formula Regional European Championship", "badge": "FRECA", "color_cat": "freca", "excel_note": "Theme 5 (Orange): Formula Regional Europe (Alpine)", "title": "FRECA Silverstone", "circuit": "Silverstone Circuit", "city": "Silverstone", "country": "United Kingdom", "flag": "🇬🇧", "notes": "", "theme": 5, "rgb": null}]}, {"week_num": 21, "excel_row": 23, "start_date": "2027-05-21", "end_date": "2027-05-23", "date_formatted": "May 21 – 23, 2027", "month": "May", "month_idx": 5, "quarter": "Q2", "status": "Multi-Series Weekend", "status_desc": "⚡ 3 Championships in Action", "event_count": 3, "events": [{"col": 3, "row": 23, "raw_text": "Canada (s)", "series": "Formula 1", "sub_series": "F1 • F2 & F3 Support", "badge": "F1 + F2 + F3 • Sprint", "color_cat": "f1-f3", "excel_note": "Green Cell: F3 & F2 Support Races Included • Sprint Weekend (s)", "title": "Formula 1 Canadian Grand Prix", "circuit": "Circuit Gilles Villeneuve", "city": "Montreal", "country": "Canada", "flag": "🇨🇦", "notes": "F1 Sprint Weekend • Wall of Champions", "theme": 6, "rgb": null}, {"col": 6, "row": 23, "raw_text": "Imola 3 Hour", "series": "GTWCE", "sub_series": "GT World Challenge Europe", "badge": "GTWCE Enduro", "color_cat": "gtwce", "excel_note": "Theme 5 (Orange): GTWCE Endurance Cup", "title": "GTWCE Imola 3 Hour", "circuit": "Autodromo Enzo e Dino Ferrari", "city": "Imola", "country": "Italy", "flag": "🇮🇹", "notes": "GTWCE Endurance Cup", "theme": 5, "rgb": null}, {"col": 15, "row": 23, "raw_text": "Ricardo Tormo", "series": "Spanish F4", "sub_series": "F4 Spanish Championship", "badge": "Spanish F4", "color_cat": "f4-es", "excel_note": "Light Yellow: Campeonato de España de F4", "title": "Spanish F4 - Valencia", "circuit": "Circuit Ricardo Tormo", "city": "Cheste, Valencia", "country": "Spain", "flag": "🇪🇸", "notes": "", "theme": null, "rgb": "FFFFFFCC"}]}, {"week_num": 22, "excel_row": 24, "start_date": "2027-05-28", "end_date": "2027-05-30", "date_formatted": "May 28 – 30, 2027", "month": "May", "month_idx": 5, "quarter": "Q2", "status": "Race Weekend", "status_desc": "🏁 2 Championship Events", "event_count": 2, "events": [{"col": 4, "row": 24, "raw_text": "London*2", "series": "Formula E", "sub_series": "FE World Championship", "badge": "Double Header (*2)", "color_cat": "fe", "excel_note": "Double Header: 2 E-Prix Races", "title": "London E-Prix (Double Header)", "circuit": "ExCeL London Circuit", "city": "London", "country": "United Kingdom", "flag": "🇬🇧", "notes": "Rounds 1 & 2 • Indoor/Outdoor Track", "theme": null, "rgb": "00000000"}, {"col": 6, "row": 24, "raw_text": "24H Nuerburgring", "series": "IGTC", "sub_series": "Intercontinental GT Challenge", "badge": "IGTC", "color_cat": "igtc", "excel_note": "Theme 3 (Grey): Intercontinental GT Classic", "title": "55th ADAC RAVENOL 24h Nürburgring", "circuit": "Nürburgring Nordschleife", "city": "Nürburg", "country": "Germany", "flag": "🇩🇪", "notes": "The Green Hell 24h • IGTC Round", "theme": 3, "rgb": null}]}, {"week_num": 23, "excel_row": 25, "start_date": "2027-06-04", "end_date": "2027-06-06", "date_formatted": "Jun 04 – 06, 2027", "month": "June", "month_idx": 6, "quarter": "Q2", "status": "Multi-Series Weekend", "status_desc": "⚡ 3 Championships in Action", "event_count": 3, "events": [{"col": 3, "row": 25, "raw_text": "Monaco (s)", "series": "Formula 1", "sub_series": "F1 • F2 & F3 Support", "badge": "F1 + F2 + F3 • Sprint", "color_cat": "f1-f3", "excel_note": "Green Cell: F3 & F2 Support Races Included • Sprint Weekend (s)", "title": "Formula 1 Grand Prix de Monaco", "circuit": "Circuit de Monaco", "city": "Monte Carlo", "country": "Monaco", "flag": "🇲🇨", "notes": "F1 Sprint Weekend • The Jewel of Motorsport", "theme": 6, "rgb": null}, {"col": 10, "row": 25, "raw_text": "Detroit", "series": "IMSA", "sub_series": "IMSA SportsCar Championship", "badge": "IMSA", "color_cat": "imsa", "excel_note": "IMSA WeatherTech SportsCar Championship", "title": "Chevrolet Detroit Sports Car Classic", "circuit": "Streets of Downtown Detroit", "city": "Detroit, MI", "country": "United States", "flag": "🇺🇸", "notes": "100-Minute Motor City Street Sprint", "theme": null, "rgb": "00000000"}, {"col": 12, "row": 25, "raw_text": "Japan", "series": "WRC", "sub_series": "World Rally Championship", "badge": "WRC", "color_cat": "wrc", "excel_note": "FIA World Rally Championship", "title": "WRC FORUM8 Rally Japan", "circuit": "Aichi & Gifu Mountain Roads", "city": "Toyota City, Aichi", "country": "Japan", "flag": "🇯🇵", "notes": "Narrow Technical Tarmac Tunnels", "theme": null, "rgb": "00000000"}]}, {"week_num": 24, "excel_row": 26, "start_date": "2027-06-11", "end_date": "2027-06-13", "date_formatted": "Jun 11 – 13, 2027", "month": "June", "month_idx": 6, "quarter": "Q2", "status": "Race Weekend", "status_desc": "🏁 2 Championship Events", "event_count": 2, "events": [{"col": 5, "row": 26, "raw_text": "24 Hours of Le Mans", "series": "FIA WEC", "sub_series": "World Endurance Championship", "badge": "WEC", "color_cat": "wec", "excel_note": "Theme 3 (Grey): FIA World Endurance Championship", "title": "95th 24 Hours of Le Mans", "circuit": "Circuit de la Sarthe", "city": "Le Mans", "country": "France", "flag": "🇫🇷", "notes": "Crown Jewel of Global Endurance • Triple Crown Classic", "theme": 3, "rgb": null}, {"col": 6, "row": 26, "raw_text": "Suzuka", "series": "GTWC Asia", "sub_series": "GT World Challenge Asia", "badge": "GTWC Asia", "color_cat": "gtwca", "excel_note": "Theme 9 (Green): GT World Challenge Asia", "title": "GTWC Asia Suzuka", "circuit": "Suzuka International Racing Course", "city": "Suzuka", "country": "Japan", "flag": "🇯🇵", "notes": "", "theme": 9, "rgb": null}]}, {"week_num": 25, "excel_row": 27, "start_date": "2027-06-18", "end_date": "2027-06-20", "date_formatted": "Jun 18 – 20, 2027", "month": "June", "month_idx": 6, "quarter": "Q2", "status": "Mega Super Weekend", "status_desc": "🔥 6 Major Championships Racing!", "event_count": 6, "events": [{"col": 3, "row": 27, "raw_text": "Portugal", "series": "Formula 1", "sub_series": "F1 Standalone", "badge": "F1 Exclusive", "color_cat": "f1-only", "excel_note": "Uncolored: F1 Headline Only", "title": "Formula 1 Portuguese Grand Prix", "circuit": "Autódromo Internacional do Algarve", "city": "Portimão", "country": "Portugal", "flag": "🇵🇹", "notes": "Rollercoaster Elevations", "theme": null, "rgb": "00000000"}, {"col": 4, "row": 27, "raw_text": "Zandvoort*2", "series": "Formula E", "sub_series": "FE World Championship", "badge": "Double Header (*2)", "color_cat": "fe", "excel_note": "Double Header: 2 E-Prix Races", "title": "Zandvoort E-Prix (Double Header)", "circuit": "Circuit Zandvoort", "city": "Zandvoort", "country": "Netherlands", "flag": "🇳🇱", "notes": "Rounds 1 & 2 • Dunes Circuit", "theme": null, "rgb": "00000000"}, {"col": 11, "row": 27, "raw_text": "Lausitzring", "series": "DTM", "sub_series": "Deutsche Tourenwagen Masters", "badge": "DTM", "color_cat": "dtm", "excel_note": "Deutsche Tourenwagen Masters (2 Races)", "title": "DTM Lausitzring", "circuit": "DEKRA Lausitzring", "city": "Klettwitz", "country": "Germany", "flag": "🇩🇪", "notes": "Races 1 & 2 • Banked Turn 1", "theme": null, "rgb": "00000000"}, {"col": 13, "row": 27, "raw_text": "Belgium", "series": "FRECA", "sub_series": "Formula Regional European Championship", "badge": "FRECA", "color_cat": "freca", "excel_note": "Theme 5 (Orange): Formula Regional Europe (Alpine)", "title": "FRECA Spa-Francorchamps", "circuit": "Circuit de Spa-Francorchamps", "city": "Stavelot", "country": "Belgium", "flag": "🇧🇪", "notes": "", "theme": 5, "rgb": null}, {"col": 15, "row": 27, "raw_text": "Jerez", "series": "Spanish F4", "sub_series": "F4 Spanish Championship", "badge": "Spanish F4", "color_cat": "f4-es", "excel_note": "Light Yellow: Campeonato de España de F4", "title": "Spanish F4 - Jerez", "circuit": "Circuito de Jerez-Ángel Nieto", "city": "Jerez de la Frontera", "country": "Spain", "flag": "🇪🇸", "notes": "", "theme": null, "rgb": "FFFFFFCC"}, {"col": 17, "row": 27, "raw_text": "Lausitzring", "series": "German F4", "sub_series": "German F4 / ADAC", "badge": "German F4", "color_cat": "f4-de", "excel_note": "Theme 3 (Grey): ADAC Formula 4 (DTM Support)", "title": "German F4 - Lausitzring", "circuit": "DEKRA Lausitzring", "city": "Klettwitz", "country": "Germany", "flag": "🇩🇪", "notes": "DTM Support Weekend", "theme": 3, "rgb": null}]}, {"week_num": 26, "excel_row": 28, "start_date": "2027-06-25", "end_date": "2027-06-27", "date_formatted": "Jun 25 – 27, 2027", "month": "June", "month_idx": 6, "quarter": "Q2", "status": "Mega Super Weekend", "status_desc": "🔥 5 Major Championships Racing!", "event_count": 5, "events": [{"col": 4, "row": 28, "raw_text": "Madrid*2", "series": "Formula E", "sub_series": "FE World Championship", "badge": "Double Header (*2)", "color_cat": "fe", "excel_note": "Double Header: 2 E-Prix Races", "title": "Madrid E-Prix (Double Header)", "circuit": "Circuito del Jarama / Madrid", "city": "Madrid", "country": "Spain", "flag": "🇪🇸", "notes": "Rounds 1 & 2", "theme": null, "rgb": "00000000"}, {"col": 6, "row": 28, "raw_text": "24 Hours of SPA", "series": "IGTC", "sub_series": "Intercontinental GT Challenge", "badge": "IGTC", "color_cat": "igtc", "excel_note": "Theme 3 (Grey): Intercontinental GT Classic", "title": "CrowdStrike 24 Hours of Spa", "circuit": "Circuit de Spa-Francorchamps", "city": "Stavelot", "country": "Belgium", "flag": "🇧🇪", "notes": "Centenary GT3 Classic • IGTC & GTWCE Flagship", "theme": 3, "rgb": null}, {"col": 10, "row": 28, "raw_text": "Watkins Glen", "series": "IMSA", "sub_series": "IMSA SportsCar Championship", "badge": "IMSA", "color_cat": "imsa", "excel_note": "IMSA WeatherTech SportsCar Championship", "title": "Sahlen's Six Hours of The Glen", "circuit": "Watkins Glen International", "city": "Watkins Glen, NY", "country": "United States", "flag": "🇺🇸", "notes": "Michelin Endurance Cup 6h", "theme": null, "rgb": "00000000"}, {"col": 12, "row": 28, "raw_text": "Greece", "series": "WRC", "sub_series": "World Rally Championship", "badge": "WRC", "color_cat": "wrc", "excel_note": "FIA World Rally Championship", "title": "WRC EKO Acropolis Rally Greece", "circuit": "Rocky Mountain Gravel Stages", "city": "Lamia", "country": "Greece", "flag": "🇬🇷", "notes": "Rally of Gods • High Heat & Punishing Boulders", "theme": null, "rgb": "00000000"}, {"col": 15, "row": 28, "raw_text": "Vallelunga", "series": "Italian F4", "sub_series": "Italian F4 / Euro 4", "badge": "Italian F4", "color_cat": "f4-it", "excel_note": "Theme 6 (Green): Italian F4 / Euro 4 Championship", "title": "Italian F4 - Vallelunga", "circuit": "Autodromo Vallelunga Piero Taruffi", "city": "Campagnano di Roma", "country": "Italy", "flag": "🇮🇹", "notes": "", "theme": 6, "rgb": null}]}, {"week_num": 27, "excel_row": 29, "start_date": "2027-07-02", "end_date": "2027-07-04", "date_formatted": "Jul 02 – 04, 2027", "month": "July", "month_idx": 7, "quarter": "Q3", "status": "Multi-Series Weekend", "status_desc": "⚡ 4 Championships in Action", "event_count": 4, "events": [{"col": 3, "row": 29, "raw_text": "Great Britian (s)", "series": "Formula 1", "sub_series": "F1 • F2 & F3 Support", "badge": "F1 + F2 + F3 • Sprint", "color_cat": "f1-f3", "excel_note": "Green Cell: F3 & F2 Support Races Included • Sprint Weekend (s)", "title": "Formula 1 British Grand Prix", "circuit": "Silverstone Circuit", "city": "Silverstone, Northamptonshire", "country": "United Kingdom", "flag": "🇬🇧", "notes": "F1 Sprint Weekend • Maggots, Becketts & Chapel", "theme": 6, "rgb": null}, {"col": 5, "row": 29, "raw_text": "Imola", "series": "ELMS", "sub_series": "European Le Mans Series", "badge": "ELMS", "color_cat": "elms", "excel_note": "Theme 5 (Orange): European Le Mans Series", "title": "ELMS 4 Hours of Imola", "circuit": "Autodromo Enzo e Dino Ferrari", "city": "Imola", "country": "Italy", "flag": "🇮🇹", "notes": "", "theme": 5, "rgb": null}, {"col": 11, "row": 29, "raw_text": "Norisring", "series": "DTM", "sub_series": "Deutsche Tourenwagen Masters", "badge": "DTM", "color_cat": "dtm", "excel_note": "Deutsche Tourenwagen Masters (2 Races)", "title": "DTM Norisring", "circuit": "Norisring", "city": "", "country": "", "flag": "🏁", "notes": "", "theme": null, "rgb": "00000000"}, {"col": 15, "row": 29, "raw_text": "Norisring", "series": "German F4", "sub_series": "German F4 / ADAC", "badge": "German F4", "color_cat": "f4-de", "excel_note": "Theme 3 (Grey): ADAC Formula 4 (DTM Support)", "title": "German F4 Norisring", "circuit": "Norisring", "city": "", "country": "", "flag": "🏁", "notes": "", "theme": 3, "rgb": null}]}, {"week_num": 28, "excel_row": 30, "start_date": "2027-07-09", "end_date": "2027-07-11", "date_formatted": "Jul 09 – 11, 2027", "month": "July", "month_idx": 7, "quarter": "Q3", "status": "Mega Super Weekend", "status_desc": "🔥 6 Major Championships Racing!", "event_count": 6, "events": [{"col": 3, "row": 30, "raw_text": "Austria", "series": "Formula 1", "sub_series": "F1 • F2 & F3 Support", "badge": "F1 + F2 + F3", "color_cat": "f1-f3", "excel_note": "Green Cell: F3 & F2 Support Races Included", "title": "Formula 1 Austrian Grand Prix", "circuit": "Red Bull Ring", "city": "Spielberg, Styria", "country": "Austria", "flag": "🇦🇹", "notes": "Styrian Alpine Backdrop", "theme": 6, "rgb": null}, {"col": 4, "row": 30, "raw_text": "Shanghai*2", "series": "Formula E", "sub_series": "FE World Championship", "badge": "Double Header (*2)", "color_cat": "fe", "excel_note": "Double Header: 2 E-Prix Races", "title": "Shanghai E-Prix (Double Header)", "circuit": "Shanghai International Circuit", "city": "Shanghai", "country": "China", "flag": "🇨🇳", "notes": "Rounds 1 & 2", "theme": null, "rgb": "00000000"}, {"col": 5, "row": 30, "raw_text": "6 Hours of São Paulo", "series": "FIA WEC", "sub_series": "World Endurance Championship", "badge": "WEC", "color_cat": "wec", "excel_note": "Theme 3 (Grey): FIA World Endurance Championship", "title": "FIA WEC 6 Hours of São Paulo", "circuit": "Autódromo José Carlos Pace (Interlagos)", "city": "São Paulo", "country": "Brazil", "flag": "🇧🇷", "notes": "", "theme": 3, "rgb": null}, {"col": 6, "row": 30, "raw_text": "Fuji", "series": "GTWC Asia", "sub_series": "GT World Challenge Asia", "badge": "GTWC Asia", "color_cat": "gtwca", "excel_note": "Theme 9 (Green): GT World Challenge Asia", "title": "GTWC Asia Fuji", "circuit": "Fuji International Speedway", "city": "Oyama", "country": "Japan", "flag": "🇯🇵", "notes": "", "theme": 9, "rgb": null}, {"col": 10, "row": 30, "raw_text": "Canadian Tire \nMotorsport Park", "series": "IMSA", "sub_series": "IMSA SportsCar Championship", "badge": "IMSA", "color_cat": "imsa", "excel_note": "IMSA WeatherTech SportsCar Championship", "title": "Chevrolet Grand Prix at Mosport", "circuit": "Canadian Tire Motorsport Park", "city": "Bowmanville, Ontario", "country": "Canada", "flag": "🇨🇦", "notes": "Ultra High-Speed Road Course • 2h 40m", "theme": null, "rgb": "00000000"}, {"col": 12, "row": 30, "raw_text": "Italy", "series": "WRC", "sub_series": "World Rally Championship", "badge": "WRC", "color_cat": "wrc", "excel_note": "FIA World Rally Championship", "title": "WRC Rally Italia Sardegna", "circuit": "Mediterranean Island Gravel", "city": "Olbia / Alghero, Sardinia", "country": "Italy", "flag": "🇮🇹", "notes": "Sandy Gravel, Water Splashes & Micky's Jump", "theme": null, "rgb": "00000000"}]}, {"week_num": 29, "excel_row": 31, "start_date": "2027-07-16", "end_date": "2027-07-18", "date_formatted": "Jul 16 – 18, 2027", "month": "July", "month_idx": 7, "quarter": "Q3", "status": "Multi-Series Weekend", "status_desc": "⚡ 3 Championships in Action", "event_count": 3, "events": [{"col": 6, "row": 31, "raw_text": "Misano (s)", "series": "GTWCE", "sub_series": "GT World Challenge Europe", "badge": "GTWCE Sprint (s)", "color_cat": "gtwce", "excel_note": "Theme 5 (Orange): GTWCE Sprint Cup (s)", "title": "GTWCE Misano Sprint Cup", "circuit": "Misano World Circuit Marco Simoncelli", "city": "Misano Adriatico", "country": "Italy", "flag": "🇮🇹", "notes": "Night Sprint Races", "theme": 5, "rgb": null}, {"col": 13, "row": 31, "raw_text": "France", "series": "FRECA", "sub_series": "Formula Regional European Championship", "badge": "FRECA", "color_cat": "freca", "excel_note": "Theme 5 (Orange): Formula Regional Europe (Alpine)", "title": "FRECA Paul Ricard", "circuit": "Circuit Paul Ricard", "city": "Le Castellet", "country": "France", "flag": "🇫🇷", "notes": "", "theme": 5, "rgb": null}, {"col": 15, "row": 31, "raw_text": "Algarve", "series": "Spanish F4", "sub_series": "F4 Spanish Championship", "badge": "Spanish F4", "color_cat": "f4-es", "excel_note": "Light Yellow: Campeonato de España de F4", "title": "Spanish F4 - Portimão", "circuit": "Autódromo Internacional do Algarve", "city": "Portimão", "country": "Portugal", "flag": "🇵🇹", "notes": "", "theme": null, "rgb": "FFFFFFCC"}]}, {"week_num": 30, "excel_row": 32, "start_date": "2027-07-23", "end_date": "2027-07-25", "date_formatted": "Jul 23 – 25, 2027", "month": "July", "month_idx": 7, "quarter": "Q3", "status": "Multi-Series Weekend", "status_desc": "⚡ 4 Championships in Action", "event_count": 4, "events": [{"col": 3, "row": 32, "raw_text": "Belgium", "series": "Formula 1", "sub_series": "F1 Standalone", "badge": "F1 Exclusive", "color_cat": "f1-only", "excel_note": "Uncolored: F1 Headline Only", "title": "Formula 1 Belgian Grand Prix", "circuit": "Circuit de Spa-Francorchamps", "city": "Stavelot", "country": "Belgium", "flag": "🇧🇪", "notes": "Eau Rouge, Raidillon & Pouhon", "theme": null, "rgb": "00000000"}, {"col": 4, "row": 32, "raw_text": "Tokyo*2", "series": "Formula E", "sub_series": "FE World Championship", "badge": "Double Header (*2)", "color_cat": "fe", "excel_note": "Double Header: 2 E-Prix Races", "title": "Tokyo E-Prix (Double Header)", "circuit": "Tokyo Big Sight Street Circuit", "city": "Tokyo", "country": "Japan", "flag": "🇯🇵", "notes": "Rounds 1 & 2 • Waterfront Docks", "theme": null, "rgb": "00000000"}, {"col": 12, "row": 32, "raw_text": "Estonia", "series": "WRC", "sub_series": "World Rally Championship", "badge": "WRC", "color_cat": "wrc", "excel_note": "FIA World Rally Championship", "title": "WRC Rally Estonia", "circuit": "High-Speed Rollercoaster Gravel", "city": "Tartu", "country": "Estonia", "flag": "🇪🇪", "notes": "Massive High-Speed Jumps & Crests", "theme": null, "rgb": "00000000"}, {"col": 15, "row": 32, "raw_text": "Mugello", "series": "Italian F4", "sub_series": "Italian F4 / Euro 4", "badge": "Italian F4", "color_cat": "f4-it", "excel_note": "Theme 6 (Green): Italian F4 / Euro 4 Championship", "title": "Italian F4 - Mugello", "circuit": "Autodromo Internazionale del Mugello", "city": "Scarperia e San Piero", "country": "Italy", "flag": "🇮🇹", "notes": "", "theme": 6, "rgb": null}]}, {"week_num": 31, "excel_row": 33, "start_date": "2027-07-30", "end_date": "2027-08-01", "date_formatted": "Jul 30 – Aug 01, 2027", "month": "July", "month_idx": 7, "quarter": "Q3", "status": "Mega Super Weekend", "status_desc": "🔥 6 Major Championships Racing!", "event_count": 6, "events": [{"col": 3, "row": 33, "raw_text": "Hungary", "series": "Formula 1", "sub_series": "F1 Standalone", "badge": "F1 Exclusive", "color_cat": "f1-only", "excel_note": "Uncolored: F1 Headline Only", "title": "Formula 1 Hungarian Grand Prix", "circuit": "Hungaroring", "city": "Mogyoród", "country": "Hungary", "flag": "🇭🇺", "notes": "Monaco without the Buildings", "theme": null, "rgb": "00000000"}, {"col": 6, "row": 33, "raw_text": "Magny-cours (s)", "series": "GTWCE", "sub_series": "GT World Challenge Europe", "badge": "GTWCE Sprint (s)", "color_cat": "gtwce", "excel_note": "Theme 5 (Orange): GTWCE Sprint Cup (s)", "title": "GTWCE Magny-Cours Sprint Cup", "circuit": "Circuit de Nevers Magny-Cours", "city": "Magny-Cours", "country": "France", "flag": "🇫🇷", "notes": "GTWCE Sprint Cup", "theme": 5, "rgb": null}, {"col": 12, "row": 33, "raw_text": "Finland", "series": "WRC", "sub_series": "World Rally Championship", "badge": "WRC", "color_cat": "wrc", "excel_note": "FIA World Rally Championship", "title": "WRC Secto Rally Finland", "circuit": "Gravel Grand Prix of Forest Jumps", "city": "Jyväskylä", "country": "Finland", "flag": "🇫🇮", "notes": "Fastest Rally in the World • 130+ km/h avg", "theme": null, "rgb": "00000000"}, {"col": 13, "row": 33, "raw_text": "Hungary", "series": "FRECA", "sub_series": "Formula Regional European Championship", "badge": "FRECA", "color_cat": "freca", "excel_note": "Theme 5 (Orange): Formula Regional Europe (Alpine)", "title": "FRECA Hungaroring", "circuit": "Hungaroring", "city": "Mogyoród", "country": "Hungary", "flag": "🇭🇺", "notes": "", "theme": 5, "rgb": null}, {"col": 15, "row": 33, "raw_text": "Aragon", "series": "Spanish F4", "sub_series": "F4 Spanish Championship", "badge": "Spanish F4", "color_cat": "f4-es", "excel_note": "Light Yellow: Campeonato de España de F4", "title": "Spanish F4 - MotorLand Aragón", "circuit": "MotorLand Aragón", "city": "Alcañiz, Teruel", "country": "Spain", "flag": "🇪🇸", "notes": "", "theme": null, "rgb": "FFFFFFCC"}, {"col": 17, "row": 33, "raw_text": "Red Bull Ring", "series": "German F4", "sub_series": "German F4 / ADAC", "badge": "German F4", "color_cat": "f4-de", "excel_note": "Theme 3 (Grey): ADAC Formula 4 (DTM Support)", "title": "German F4 - Red Bull Ring", "circuit": "Red Bull Ring", "city": "Spielberg", "country": "Austria", "flag": "🇦🇹", "notes": "ADAC Championship Round", "theme": 3, "rgb": null}]}, {"week_num": 32, "excel_row": 34, "start_date": "2027-08-06", "end_date": "2027-08-08", "date_formatted": "Aug 06 – 08, 2027", "month": "August", "month_idx": 8, "quarter": "Q3", "status": "Multi-Series Weekend", "status_desc": "⚡ 4 Championships in Action", "event_count": 4, "events": [{"col": 10, "row": 34, "raw_text": "Road America", "series": "IMSA", "sub_series": "IMSA SportsCar Championship", "badge": "IMSA", "color_cat": "imsa", "excel_note": "IMSA WeatherTech SportsCar Championship", "title": "IMSA SportsCar Weekend at Road America", "circuit": "Road America", "city": "Elkhart Lake, WI", "country": "United States", "flag": "🇺🇸", "notes": "4.048-mile Circuit • 2h 40m", "theme": null, "rgb": "00000000"}, {"col": 11, "row": 34, "raw_text": "Zandvoort", "series": "DTM", "sub_series": "Deutsche Tourenwagen Masters", "badge": "DTM", "color_cat": "dtm", "excel_note": "Deutsche Tourenwagen Masters (2 Races)", "title": "DTM Circuit Zandvoort", "circuit": "Circuit Zandvoort", "city": "Zandvoort", "country": "Netherlands", "flag": "🇳🇱", "notes": "Races 1 & 2 • Coastal Dunes", "theme": null, "rgb": "00000000"}, {"col": 13, "row": 34, "raw_text": "Netherlands", "series": "FRECA", "sub_series": "Formula Regional European Championship", "badge": "FRECA", "color_cat": "freca", "excel_note": "Theme 5 (Orange): Formula Regional Europe (Alpine)", "title": "FRECA Zandvoort", "circuit": "Circuit Zandvoort", "city": "Zandvoort", "country": "Netherlands", "flag": "🇳🇱", "notes": "", "theme": 5, "rgb": null}, {"col": 15, "row": 34, "raw_text": "Misano", "series": "Italian F4", "sub_series": "Italian F4 / Euro 4", "badge": "Italian F4", "color_cat": "f4-it", "excel_note": "Theme 6 (Green): Italian F4 / Euro 4 Championship", "title": "Italian F4 - Misano", "circuit": "Misano World Circuit Marco Simoncelli", "city": "Misano Adriatico", "country": "Italy", "flag": "🇮🇹", "notes": "", "theme": 6, "rgb": null}]}, {"week_num": 33, "excel_row": 35, "start_date": "2027-08-13", "end_date": "2027-08-15", "date_formatted": "Aug 13 – 15, 2027", "month": "August", "month_idx": 8, "quarter": "Q3", "status": "Summer Break", "status_desc": "Mid-Season Summer Hiatus", "event_count": 0, "events": []}, {"week_num": 34, "excel_row": 36, "start_date": "2027-08-20", "end_date": "2027-08-22", "date_formatted": "Aug 20 – 22, 2027", "month": "August", "month_idx": 8, "quarter": "Q3", "status": "Multi-Series Weekend", "status_desc": "⚡ 4 Championships in Action", "event_count": 4, "events": [{"col": 5, "row": 36, "raw_text": "Spa-Francorchamps", "series": "ELMS", "sub_series": "European Le Mans Series", "badge": "ELMS", "color_cat": "elms", "excel_note": "Theme 5 (Orange): European Le Mans Series", "title": "ELMS 4 Hours of Spa-Francorchamps", "circuit": "Circuit de Spa-Francorchamps", "city": "Stavelot", "country": "Belgium", "flag": "🇧🇪", "notes": "", "theme": 5, "rgb": null}, {"col": 10, "row": 36, "raw_text": "Virginia", "series": "IMSA", "sub_series": "IMSA SportsCar Championship", "badge": "IMSA", "color_cat": "imsa", "excel_note": "IMSA WeatherTech SportsCar Championship", "title": "Michelin GT Challenge at VIR", "circuit": "Virginia International Raceway", "city": "Alton, VA", "country": "United States", "flag": "🇺🇸", "notes": "GTD PRO / GTD Headline Showcase • 2h 40m", "theme": null, "rgb": "00000000"}, {"col": 11, "row": 36, "raw_text": "Nuerburgring", "series": "DTM", "sub_series": "Deutsche Tourenwagen Masters", "badge": "DTM", "color_cat": "dtm", "excel_note": "Deutsche Tourenwagen Masters (2 Races)", "title": "DTM Nuerburgring", "circuit": "Nuerburgring", "city": "", "country": "", "flag": "🏁", "notes": "", "theme": null, "rgb": "00000000"}, {"col": 15, "row": 36, "raw_text": "Nuerburgring", "series": "German F4", "sub_series": "German F4 / ADAC", "badge": "German F4", "color_cat": "f4-de", "excel_note": "Theme 3 (Grey): ADAC Formula 4 (DTM Support)", "title": "German F4 - Nürburgring", "circuit": "Nürburgring GP Circuit", "city": "Nürburg", "country": "Germany", "flag": "🇩🇪", "notes": "DTM Support Weekend", "theme": 3, "rgb": null}]}, {"week_num": 35, "excel_row": 37, "start_date": "2027-08-27", "end_date": "2027-08-29", "date_formatted": "Aug 27 – 29, 2027", "month": "August", "month_idx": 8, "quarter": "Q3", "status": "Race Weekend", "status_desc": "🏁 2 Championship Events", "event_count": 2, "events": [{"col": 6, "row": 37, "raw_text": "Nuerburgring 3 Hour", "series": "GTWCE", "sub_series": "GT World Challenge Europe", "badge": "GTWCE Enduro", "color_cat": "gtwce", "excel_note": "Theme 5 (Orange): GTWCE Endurance Cup", "title": "GTWCE Nürburgring 3 Hour", "circuit": "Nürburgring GP Circuit", "city": "Nürburg", "country": "Germany", "flag": "🇩🇪", "notes": "GTWCE Endurance Cup", "theme": 5, "rgb": null}, {"col": 12, "row": 37, "raw_text": "United States", "series": "WRC", "sub_series": "World Rally Championship", "badge": "WRC", "color_cat": "wrc", "excel_note": "FIA World Rally Championship", "title": "WRC Rally USA (Chattanooga)", "circuit": "Cherokee National Forest Gravel Stages", "city": "Chattanooga, TN", "country": "United States", "flag": "🇺🇸", "notes": "Historic American WRC Return", "theme": null, "rgb": "00000000"}]}, {"week_num": 36, "excel_row": 38, "start_date": "2027-09-03", "end_date": "2027-09-05", "date_formatted": "Sep 03 – 05, 2027", "month": "September", "month_idx": 9, "quarter": "Q3", "status": "Multi-Series Weekend", "status_desc": "⚡ 3 Championships in Action", "event_count": 3, "events": [{"col": 3, "row": 38, "raw_text": "Italy (s)", "series": "Formula 1", "sub_series": "F1 • F2 & F3 Support", "badge": "F1 + F2 + F3 • Sprint", "color_cat": "f1-f3", "excel_note": "Green Cell: F3 & F2 Support Races Included • Sprint Weekend (s)", "title": "Formula 1 Pirelli Gran Premio d'Italia", "circuit": "Autodromo Nazionale Monza", "city": "Monza, Lombardy", "country": "Italy", "flag": "🇮🇹", "notes": "F1 Sprint Weekend • Temple of Speed", "theme": 6, "rgb": null}, {"col": 5, "row": 38, "raw_text": "Silverstone", "series": "ELMS", "sub_series": "European Le Mans Series", "badge": "ELMS", "color_cat": "elms", "excel_note": "Theme 5 (Orange): European Le Mans Series", "title": "ELMS 4 Hours of Silverstone", "circuit": "Silverstone Circuit", "city": "Silverstone", "country": "United Kingdom", "flag": "🇬🇧", "notes": "", "theme": 5, "rgb": null}, {"col": 6, "row": 38, "raw_text": "Suzuka 1000km", "series": "IGTC", "sub_series": "Intercontinental GT Challenge", "badge": "IGTC", "color_cat": "igtc", "excel_note": "Theme 3 (Grey): Intercontinental GT Classic", "title": "IGTC Suzuka 1000km", "circuit": "Suzuka International Racing Course", "city": "Suzuka", "country": "Japan", "flag": "🇯🇵", "notes": "Legendary Summer Endurance Classic • IGTC", "theme": 3, "rgb": null}]}, {"week_num": 37, "excel_row": 39, "start_date": "2027-09-10", "end_date": "2027-09-12", "date_formatted": "Sep 10 – 12, 2027", "month": "September", "month_idx": 9, "quarter": "Q3", "status": "Multi-Series Weekend", "status_desc": "⚡ 4 Championships in Action", "event_count": 4, "events": [{"col": 3, "row": 39, "raw_text": "Spain", "series": "Formula 1", "sub_series": "F1 • F2 & F3 Support", "badge": "F1 + F2 + F3", "color_cat": "f1-f3", "excel_note": "Green Cell: F3 & F2 Support Races Included", "title": "Formula 1 Spanish Grand Prix", "circuit": "Circuit de Barcelona-Catalunya", "city": "Montmeló, Catalonia", "country": "Spain", "flag": "🇪🇸", "notes": "Testing Benchmark Circuit", "theme": 6, "rgb": null}, {"col": 5, "row": 39, "raw_text": "Lone Star Le Mans", "series": "FIA WEC", "sub_series": "World Endurance Championship", "badge": "WEC", "color_cat": "wec", "excel_note": "Theme 3 (Grey): FIA World Endurance Championship", "title": "FIA WEC Lone Star Le Mans (6 Hours)", "circuit": "Circuit of the Americas (COTA)", "city": "Austin, TX", "country": "United States", "flag": "🇺🇸", "notes": "", "theme": 3, "rgb": null}, {"col": 13, "row": 39, "raw_text": "Imola", "series": "FRECA", "sub_series": "Formula Regional European Championship", "badge": "FRECA", "color_cat": "freca", "excel_note": "Theme 5 (Orange): Formula Regional Europe (Alpine)", "title": "FRECA Imola", "circuit": "Autodromo Enzo e Dino Ferrari", "city": "Imola", "country": "Italy", "flag": "🇮🇹", "notes": "", "theme": 5, "rgb": null}, {"col": 15, "row": 39, "raw_text": "Imola", "series": "Italian F4", "sub_series": "Italian F4 / Euro 4", "badge": "Italian F4", "color_cat": "f4-it", "excel_note": "Theme 6 (Green): Italian F4 / Euro 4 Championship", "title": "Italian F4 - Imola", "circuit": "Autodromo Enzo e Dino Ferrari", "city": "Imola", "country": "Italy", "flag": "🇮🇹", "notes": "", "theme": 6, "rgb": null}]}, {"week_num": 38, "excel_row": 40, "start_date": "2027-09-17", "end_date": "2027-09-19", "date_formatted": "Sep 17 – 19, 2027", "month": "September", "month_idx": 9, "quarter": "Q3", "status": "Multi-Series Weekend", "status_desc": "⚡ 3 Championships in Action", "event_count": 3, "events": [{"col": 6, "row": 40, "raw_text": "Zandvoort (s)", "series": "GTWCE", "sub_series": "GT World Challenge Europe", "badge": "GTWCE Sprint (s)", "color_cat": "gtwce", "excel_note": "Theme 5 (Orange): GTWCE Sprint Cup (s)", "title": "GTWCE Zandvoort Sprint Cup", "circuit": "Circuit Zandvoort", "city": "Zandvoort", "country": "Netherlands", "flag": "🇳🇱", "notes": "GTWCE Sprint Cup", "theme": 5, "rgb": null}, {"col": 10, "row": 40, "raw_text": "Indianapolis", "series": "IMSA", "sub_series": "IMSA SportsCar Championship", "badge": "IMSA", "color_cat": "imsa", "excel_note": "IMSA WeatherTech SportsCar Championship", "title": "TireRack.com Battle on the Bricks (6 Hours)", "circuit": "Indianapolis Motor Speedway Road Course", "city": "Speedway, IN", "country": "United States", "flag": "🇺🇸", "notes": "Brickyard 6h Endurance Cup", "theme": null, "rgb": "00000000"}, {"col": 15, "row": 40, "raw_text": "Misano", "series": "Italian F4", "sub_series": "Italian F4 / Euro 4", "badge": "Italian F4", "color_cat": "f4-it", "excel_note": "Theme 6 (Green): Italian F4 / Euro 4 Championship", "title": "Italian F4 - Misano", "circuit": "Misano World Circuit Marco Simoncelli", "city": "Misano Adriatico", "country": "Italy", "flag": "🇮🇹", "notes": "", "theme": 6, "rgb": null}]}, {"week_num": 39, "excel_row": 41, "start_date": "2027-09-24", "end_date": "2027-09-26", "date_formatted": "Sep 24 – 26, 2027", "month": "September", "month_idx": 9, "quarter": "Q3", "status": "Mega Super Weekend", "status_desc": "🔥 6 Major Championships Racing!", "event_count": 6, "events": [{"col": 3, "row": 41, "raw_text": "Azerbaijan", "series": "Formula 1", "sub_series": "F1 • F2 Support", "badge": "F1 + F2", "color_cat": "f1-f2", "excel_note": "Blue Cell: F2 Support Race Included", "title": "Formula 1 Azerbaijan Grand Prix", "circuit": "Baku City Circuit", "city": "Baku", "country": "Azerbaijan", "flag": "🇦🇿", "notes": "Castle Section & 2.2km Straight", "theme": 7, "rgb": null}, {"col": 5, "row": 41, "raw_text": "6 Hours of Fuji", "series": "FIA WEC", "sub_series": "World Endurance Championship", "badge": "WEC", "color_cat": "wec", "excel_note": "Theme 3 (Grey): FIA World Endurance Championship", "title": "FIA WEC 6 Hours of Fuji", "circuit": "Fuji International Speedway", "city": "Oyama, Shizuoka", "country": "Japan", "flag": "🇯🇵", "notes": "Mount Fuji Backdrop", "theme": 3, "rgb": null}, {"col": 11, "row": 41, "raw_text": "Sachsenring", "series": "DTM", "sub_series": "Deutsche Tourenwagen Masters", "badge": "DTM", "color_cat": "dtm", "excel_note": "Deutsche Tourenwagen Masters (2 Races)", "title": "DTM Sachsenring", "circuit": "Sachsenring", "city": "", "country": "", "flag": "🏁", "notes": "", "theme": null, "rgb": "00000000"}, {"col": 13, "row": 41, "raw_text": "Monza", "series": "FRECA", "sub_series": "Formula Regional European Championship", "badge": "FRECA", "color_cat": "freca", "excel_note": "Theme 5 (Orange): Formula Regional Europe (Alpine)", "title": "FRECA Monza", "circuit": "Autodromo Nazionale Monza", "city": "Monza", "country": "Italy", "flag": "🇮🇹", "notes": "", "theme": 5, "rgb": null}, {"col": 15, "row": 41, "raw_text": "Monza", "series": "Italian F4", "sub_series": "Italian F4 / Euro 4", "badge": "Italian F4", "color_cat": "f4-it", "excel_note": "Theme 6 (Green): Italian F4 / Euro 4 Championship", "title": "Italian F4 - Monza", "circuit": "Autodromo Nazionale Monza", "city": "Monza", "country": "Italy", "flag": "🇮🇹", "notes": "", "theme": 6, "rgb": null}, {"col": 17, "row": 41, "raw_text": "Sachsenring", "series": "German F4", "sub_series": "German F4 / ADAC", "badge": "German F4", "color_cat": "f4-de", "excel_note": "Theme 3 (Grey): ADAC Formula 4 (DTM Support)", "title": "German F4 Sachsenring", "circuit": "Sachsenring", "city": "", "country": "", "flag": "🏁", "notes": "", "theme": 3, "rgb": null}]}, {"week_num": 40, "excel_row": 42, "start_date": "2027-10-01", "end_date": "2027-10-03", "date_formatted": "Oct 01 – 03, 2027", "month": "October", "month_idx": 10, "quarter": "Q4", "status": "Mega Super Weekend", "status_desc": "🔥 6 Major Championships Racing!", "event_count": 6, "events": [{"col": 3, "row": 42, "raw_text": "Turkiye", "series": "Formula 1", "sub_series": "F1 • F2 & F3 Support", "badge": "F1 + F2 + F3", "color_cat": "f1-f3", "excel_note": "Green Cell: F3 & F2 Support Races Included", "title": "Formula 1 Turkish Grand Prix", "circuit": "Intercity Istanbul Park", "city": "Tuzla, Istanbul", "country": "Turkey", "flag": "🇹🇷", "notes": "Famous Quadruple-Apex Turn 8", "theme": 6, "rgb": null}, {"col": 5, "row": 42, "raw_text": "Portimao", "series": "ELMS", "sub_series": "European Le Mans Series", "badge": "ELMS", "color_cat": "elms", "excel_note": "Theme 5 (Orange): European Le Mans Series", "title": "ELMS 4 Hours of Portimão Finale", "circuit": "Autódromo Internacional do Algarve", "city": "Portimão", "country": "Portugal", "flag": "🇵🇹", "notes": "ELMS Season Finale", "theme": 5, "rgb": null}, {"col": 6, "row": 42, "raw_text": "Budapest (s)", "series": "GTWCE", "sub_series": "GT World Challenge Europe", "badge": "GTWCE Sprint (s)", "color_cat": "gtwce", "excel_note": "Theme 5 (Orange): GTWCE Sprint Cup (s)", "title": "GTWCE Budapest Sprint Cup", "circuit": "Hungaroring", "city": "Mogyoród", "country": "Hungary", "flag": "🇭🇺", "notes": "GTWCE Sprint Cup", "theme": 5, "rgb": null}, {"col": 8, "row": 42, "raw_text": "Beijing", "series": "GTWC Asia", "sub_series": "GT World Challenge Asia", "badge": "GTWC Asia", "color_cat": "gtwca", "excel_note": "Theme 9 (Green): GT World Challenge Asia", "title": "GTWC Asia Beijing", "circuit": "Beijing Street / International Circuit", "city": "Beijing", "country": "China", "flag": "🇨🇳", "notes": "", "theme": 9, "rgb": null}, {"col": 12, "row": 42, "raw_text": "United Kingdom", "series": "WRC", "sub_series": "World Rally Championship", "badge": "WRC", "color_cat": "wrc", "excel_note": "FIA World Rally Championship", "title": "WRC Rally United Kingdom", "circuit": "Classic Forest & Tarmac Stages", "city": "Aberystwyth / Wales", "country": "United Kingdom", "flag": "🇬🇧", "notes": "Legendary Mud, Fog & British Forests", "theme": null, "rgb": "00000000"}, {"col": 15, "row": 42, "raw_text": "Hockenheimring", "series": "German F4", "sub_series": "German F4 / ADAC", "badge": "German F4", "color_cat": "f4-de", "excel_note": "Theme 3 (Grey): ADAC Formula 4 (DTM Support)", "title": "German F4 Hockenheimring", "circuit": "Hockenheimring", "city": "", "country": "", "flag": "🏁", "notes": "", "theme": 3, "rgb": null}]}, {"week_num": 41, "excel_row": 43, "start_date": "2027-10-08", "end_date": "2027-10-10", "date_formatted": "Oct 08 – 10, 2027", "month": "October", "month_idx": 10, "quarter": "Q4", "status": "Mega Super Weekend", "status_desc": "🔥 5 Major Championships Racing!", "event_count": 5, "events": [{"col": 3, "row": 43, "raw_text": "Singapore", "series": "Formula 1", "sub_series": "F1 Standalone", "badge": "F1 Exclusive", "color_cat": "f1-only", "excel_note": "Uncolored: F1 Headline Only", "title": "Formula 1 Singapore Airlines Singapore Grand Prix", "circuit": "Marina Bay Street Circuit", "city": "Marina Bay", "country": "Singapore", "flag": "🇸🇬", "notes": "Original F1 Night Race Under Lights", "theme": null, "rgb": "00000000"}, {"col": 10, "row": 43, "raw_text": "Petit Le Mans", "series": "IMSA", "sub_series": "IMSA SportsCar Championship", "badge": "IMSA", "color_cat": "imsa", "excel_note": "IMSA WeatherTech SportsCar Championship", "title": "Motul Petit Le Mans (10 Hours)", "circuit": "Michelin Raceway Road Atlanta", "city": "Braselton, GA", "country": "United States", "flag": "🇺🇸", "notes": "IMSA Season Finale • 10-Hour Thriller", "theme": null, "rgb": "00000000"}, {"col": 11, "row": 43, "raw_text": "Hockenheimring", "series": "DTM", "sub_series": "Deutsche Tourenwagen Masters", "badge": "DTM", "color_cat": "dtm", "excel_note": "Deutsche Tourenwagen Masters (2 Races)", "title": "DTM Hockenheimring Season Finale", "circuit": "Hockenheimring Baden-Württemberg", "city": "Hockenheim", "country": "Germany", "flag": "🇩🇪", "notes": "DTM Championship Finale • Races 1 & 2", "theme": null, "rgb": "00000000"}, {"col": 13, "row": 43, "raw_text": "Barcelona", "series": "FRECA", "sub_series": "Formula Regional European Championship", "badge": "FRECA", "color_cat": "freca", "excel_note": "Theme 5 (Orange): Formula Regional Europe (Alpine)", "title": "FRECA Barcelona Season Finale", "circuit": "Circuit de Barcelona-Catalunya", "city": "Montmeló", "country": "Spain", "flag": "🇪🇸", "notes": "FRECA Championship Finale", "theme": 5, "rgb": null}, {"col": 15, "row": 43, "raw_text": "Navarra", "series": "Spanish F4", "sub_series": "F4 Spanish Championship", "badge": "Spanish F4", "color_cat": "f4-es", "excel_note": "Light Yellow: Campeonato de España de F4", "title": "Spanish F4 - Navarra", "circuit": "Circuito de Navarra", "city": "Los Arcos, Navarra", "country": "Spain", "flag": "🇪🇸", "notes": "", "theme": null, "rgb": "FFFFFFCC"}]}, {"week_num": 42, "excel_row": 44, "start_date": "2027-10-15", "end_date": "2027-10-17", "date_formatted": "Oct 15 – 17, 2027", "month": "October", "month_idx": 10, "quarter": "Q4", "status": "Race Weekend", "status_desc": "🏁 1 Championship Event", "event_count": 1, "events": [{"col": 15, "row": 44, "raw_text": "Barcelona", "series": "Italian F4", "sub_series": "Italian F4 / Euro 4", "badge": "Italian F4", "color_cat": "f4-it", "excel_note": "Theme 6 (Green): Italian F4 / Euro 4 Championship", "title": "Euro 4 Championship - Barcelona Finale", "circuit": "Circuit de Barcelona-Catalunya", "city": "Montmeló", "country": "Spain", "flag": "🇪🇸", "notes": "Euro 4 Championship Season Finale", "theme": 6, "rgb": null}]}, {"week_num": 43, "excel_row": 45, "start_date": "2027-10-22", "end_date": "2027-10-24", "date_formatted": "Oct 22 – 24, 2027", "month": "October", "month_idx": 10, "quarter": "Q4", "status": "Multi-Series Weekend", "status_desc": "⚡ 4 Championships in Action", "event_count": 4, "events": [{"col": 3, "row": 45, "raw_text": "United States", "series": "Formula 1", "sub_series": "F1 Standalone", "badge": "F1 Exclusive", "color_cat": "f1-only", "excel_note": "Uncolored: F1 Headline Only", "title": "Formula 1 Lenovo United States Grand Prix", "circuit": "Circuit of the Americas (COTA)", "city": "Austin, TX", "country": "United States", "flag": "🇺🇸", "notes": "Turn 1 Hillclimb & Stadium Section", "theme": null, "rgb": "00000000"}, {"col": 6, "row": 45, "raw_text": "Barcelona 3 Hour", "series": "GTWCE", "sub_series": "GT World Challenge Europe", "badge": "GTWCE Enduro", "color_cat": "gtwce", "excel_note": "Theme 5 (Orange): GTWCE Endurance Cup", "title": "GTWCE Barcelona 3 Hour Finale", "circuit": "Circuit de Barcelona-Catalunya", "city": "Montmeló", "country": "Spain", "flag": "🇪🇸", "notes": "Endurance Cup Season Finale", "theme": 5, "rgb": null}, {"col": 8, "row": 45, "raw_text": "Shanghai", "series": "GTWC Asia", "sub_series": "GT World Challenge Asia", "badge": "GTWC Asia", "color_cat": "gtwca", "excel_note": "Theme 9 (Green): GT World Challenge Asia", "title": "GTWC Asia Shanghai Finale", "circuit": "Shanghai International Circuit", "city": "Shanghai", "country": "China", "flag": "🇨🇳", "notes": "GTWC Asia Season Finale", "theme": 9, "rgb": null}, {"col": 12, "row": 45, "raw_text": "Paraguay", "series": "WRC", "sub_series": "World Rally Championship", "badge": "WRC", "color_cat": "wrc", "excel_note": "FIA World Rally Championship", "title": "WRC Rally del Paraguay", "circuit": "Itapúa Red Dirt Roads", "city": "Encarnación", "country": "Paraguay", "flag": "🇵🇾", "notes": "South American Red Dirt Debut", "theme": null, "rgb": "00000000"}]}, {"week_num": 44, "excel_row": 46, "start_date": "2027-10-29", "end_date": "2027-10-31", "date_formatted": "Oct 29 – 31, 2027", "month": "October", "month_idx": 10, "quarter": "Q4", "status": "Race Weekend", "status_desc": "🏁 2 Championship Events", "event_count": 2, "events": [{"col": 3, "row": 46, "raw_text": "Mexico City", "series": "Formula 1", "sub_series": "F1 • F2 Support", "badge": "F1 + F2", "color_cat": "f1-f2", "excel_note": "Blue Cell: F2 Support Race Included", "title": "Formula 1 Gran Premio de la Ciudad de México", "circuit": "Autódromo Hermanos Rodríguez", "city": "Mexico City", "country": "Mexico", "flag": "🇲🇽", "notes": "Foro Sol Baseball Stadium Section (2,200m altitude)", "theme": 7, "rgb": null}, {"col": 15, "row": 46, "raw_text": "Barcelona", "series": "Spanish F4", "sub_series": "F4 Spanish Championship", "badge": "Spanish F4", "color_cat": "f4-es", "excel_note": "Light Yellow: Campeonato de España de F4", "title": "Spanish F4 - Barcelona Season Finale", "circuit": "Circuit de Barcelona-Catalunya", "city": "Montmeló", "country": "Spain", "flag": "🇪🇸", "notes": "Spanish F4 Championship Finale", "theme": null, "rgb": "FFFFFFCC"}]}, {"week_num": 45, "excel_row": 47, "start_date": "2027-11-05", "end_date": "2027-11-07", "date_formatted": "Nov 05 – 07, 2027", "month": "November", "month_idx": 11, "quarter": "Q4", "status": "Multi-Series Weekend", "status_desc": "⚡ 3 Championships in Action", "event_count": 3, "events": [{"col": 3, "row": 47, "raw_text": "Sao Paulo (s)", "series": "Formula 1", "sub_series": "F1 Standalone", "badge": "F1 Exclusive • Sprint", "color_cat": "f1-only", "excel_note": "Uncolored: F1 Headline Only • Sprint Weekend (s)", "title": "Formula 1 Lenovo Grande Prêmio de São Paulo", "circuit": "Autódromo José Carlos Pace (Interlagos)", "city": "São Paulo", "country": "Brazil", "flag": "🇧🇷", "notes": "F1 Sprint Weekend • Senna S & Interlagos Bowl", "theme": null, "rgb": "00000000"}, {"col": 5, "row": 47, "raw_text": "8 Hours of Bahrain", "series": "FIA WEC", "sub_series": "World Endurance Championship", "badge": "WEC", "color_cat": "wec", "excel_note": "Theme 3 (Grey): FIA World Endurance Championship", "title": "FIA WEC 8 Hours of Bahrain Finale", "circuit": "Bahrain International Circuit", "city": "Sakhir", "country": "Bahrain", "flag": "🇧🇭", "notes": "WEC World Championship Finale • Day to Night", "theme": 3, "rgb": null}, {"col": 12, "row": 47, "raw_text": "Chile", "series": "WRC", "sub_series": "World Rally Championship", "badge": "WRC", "color_cat": "wrc", "excel_note": "FIA World Rally Championship", "title": "WRC Rally Chile BIOBÍO", "circuit": "Pacific Forest Forestry Roads", "city": "Concepción", "country": "Chile", "flag": "🇨🇱", "notes": "Fast Cambered Gravel Roads", "theme": null, "rgb": "00000000"}]}, {"week_num": 46, "excel_row": 48, "start_date": "2027-11-12", "end_date": "2027-11-14", "date_formatted": "Nov 12 – 14, 2027", "month": "November", "month_idx": 11, "quarter": "Q4", "status": "Off Week", "status_desc": "No Major Scheduled Races", "event_count": 0, "events": []}, {"week_num": 47, "excel_row": 49, "start_date": "2027-11-19", "end_date": "2027-11-21", "date_formatted": "Nov 19 – 21, 2027", "month": "November", "month_idx": 11, "quarter": "Q4", "status": "Race Weekend", "status_desc": "🏁 2 Championship Events", "event_count": 2, "events": [{"col": 3, "row": 49, "raw_text": "Las Vegas", "series": "Formula 1", "sub_series": "F1 Standalone", "badge": "F1 Exclusive", "color_cat": "f1-only", "excel_note": "Uncolored: F1 Headline Only", "title": "Formula 1 Heineken Silver Las Vegas Grand Prix", "circuit": "Las Vegas Strip Circuit", "city": "Las Vegas, NV", "country": "United States", "flag": "🇺🇸", "notes": "Saturday Night 350km/h Down The Strip", "theme": null, "rgb": "00000000"}, {"col": 13, "row": 49, "raw_text": "Macau Grand Prix", "series": "Macau GP", "sub_series": "FIA FR World Cup", "badge": "Macau GP", "color_cat": "macau", "excel_note": "Theme 0: Macau Grand Prix / FIA FR World Cup", "title": "74th Macau Grand Prix (FIA FR World Cup)", "circuit": "Guia Circuit", "city": "Macau", "country": "Macau SAR", "flag": "🇲🇴", "notes": "Historic Guia Street Circuit • World Cup", "theme": 0, "rgb": null}]}, {"week_num": 48, "excel_row": 50, "start_date": "2027-11-26", "end_date": "2027-11-28", "date_formatted": "Nov 26 – 28, 2027", "month": "November", "month_idx": 11, "quarter": "Q4", "status": "Race Weekend", "status_desc": "🏁 1 Championship Event", "event_count": 1, "events": [{"col": 12, "row": 50, "raw_text": "Saudi Arabia", "series": "WRC", "sub_series": "World Rally Championship", "badge": "WRC", "color_cat": "wrc", "excel_note": "FIA World Rally Championship", "title": "WRC Rally Saudi Arabia Finale", "circuit": "Desert Dunes, Rock & Coastal Stages", "city": "Jeddah", "country": "Saudi Arabia", "flag": "🇸🇦", "notes": "WRC World Championship Season Finale", "theme": null, "rgb": "00000000"}]}, {"week_num": 49, "excel_row": 51, "start_date": "2027-12-03", "end_date": "2027-12-05", "date_formatted": "Dec 03 – 05, 2027", "month": "December", "month_idx": 12, "quarter": "Q4", "status": "Race Weekend", "status_desc": "🏁 1 Championship Event", "event_count": 1, "events": [{"col": 3, "row": 51, "raw_text": "Qatar (s)", "series": "Formula 1", "sub_series": "F1 • F2 Support", "badge": "F1 + F2 • Sprint", "color_cat": "f1-f2", "excel_note": "Blue Cell: F2 Support Race Included • Sprint Weekend (s)", "title": "Formula 1 Qatar Airways Qatar Grand Prix", "circuit": "Lusail International Circuit", "city": "Lusail", "country": "Qatar", "flag": "🇶🇦", "notes": "F1 Sprint Weekend • High-G Night Race", "theme": 7, "rgb": null}]}, {"week_num": 50, "excel_row": 52, "start_date": "2027-12-10", "end_date": "2027-12-12", "date_formatted": "Dec 10 – 12, 2027", "month": "December", "month_idx": 12, "quarter": "Q4", "status": "Race Weekend", "status_desc": "🏁 1 Championship Event", "event_count": 1, "events": [{"col": 3, "row": 52, "raw_text": "Abu Dhabi (s)", "series": "Formula 1", "sub_series": "F1 • F2 Support", "badge": "F1 + F2 • Sprint", "color_cat": "f1-f2", "excel_note": "Blue Cell: F2 Support Race Included • Sprint Weekend (s)", "title": "Formula 1 Etihad Airways Abu Dhabi Grand Prix", "circuit": "Yas Marina Circuit", "city": "Abu Dhabi", "country": "United Arab Emirates", "flag": "🇦🇪", "notes": "F1 Sprint Weekend • Twilight Season Finale", "theme": 7, "rgb": null}]}, {"week_num": 51, "excel_row": 53, "start_date": "2027-12-17", "end_date": "2027-12-19", "date_formatted": "Dec 17 – 19, 2027", "month": "December", "month_idx": 12, "quarter": "Q4", "status": "Winter Break", "status_desc": "Off-Season Holiday Break", "event_count": 0, "events": []}, {"week_num": 52, "excel_row": 54, "start_date": "2027-12-24", "end_date": "2027-12-26", "date_formatted": "Dec 24 – 26, 2027", "month": "December", "month_idx": 12, "quarter": "Q4", "status": "Winter Break", "status_desc": "Off-Season Holiday Break", "event_count": 0, "events": []}];

    // State
    let currentWeekNum = 18; // Default to iconic Miami / Monaco weekend
    let activeView = 'spotlight';
    let selectedChampionship = 'Formula 1';

    // Color mapper for Excel cell classes
    function getCellClass(event) {
      if (!event) return '';
      if (event.col === 3) {
        if (event.theme === 6) return 'cell-f3-green';
        if (event.theme === 7) return 'cell-f2-blue';
      }
      if (event.theme === 5) return 'cell-orange';
      if (event.theme === 3) return 'cell-grey';
      if (event.theme === 9) return 'cell-lime';
      if (event.theme === 8) return 'cell-purple';
      if (event.rgb === 'FFFFFFCC') return 'cell-yellow';
      return '';
    }

    function getSeriesBadgeClass(series) {
      const map = {
        'Formula 1': 'series-f1',
        'Formula E': 'series-fe',
        'FIA WEC': 'series-wec',
        'ELMS': 'series-elms',
        'Asian Le Mans': 'series-alms',
        'IGTC': 'series-igtc',
        'GTWCE': 'series-gtwce',
        'GTWC Asia': 'series-gtwca',
        'IMSA': 'series-imsa',
        'DTM': 'series-dtm',
        'WRC': 'series-wrc',
        'FRECA': 'series-freca',
        'FROT': 'series-frot',
        'FWS': 'series-fws',
        'Italian F4': 'series-f4-it',
        'Spanish F4': 'series-f4-es',
        'German F4': 'series-f4-de',
        'Macau GP': 'series-macau'
      };
      return map[series] || 'series-igtc';
    }

    // Render Spotlight View
    function renderSpotlight(weekNum) {
      currentWeekNum = Math.max(1, Math.min(52, weekNum));
      const week = CALENDAR_DATA.find(w => w.week_num === currentWeekNum);
      if (!week) return;

      document.getElementById('currentWeekTitle').textContent = `Week ${week.week_num}`;
      document.getElementById('currentWeekDate').textContent = week.date_formatted;
      document.getElementById('weekScrubber').value = week.week_num;

      // Update Subtitle & Tag
      const heading = document.getElementById('spotlightHeading');
      const subtitle = document.getElementById('spotlightSubtitle');
      const tag = document.getElementById('spotlightIntensityTag');
      const grid = document.getElementById('spotlightEventsGrid');
      const offBanner = document.getElementById('spotlightOffWeekBanner');

      if (week.events.length === 0) {
        heading.textContent = week.status;
        subtitle.textContent = week.status_desc;
        tag.className = 'intensity-tag off';
        tag.textContent = `🌴 ${week.status}`;
        grid.style.display = 'none';
        offBanner.style.display = 'flex';
      } else {
        heading.textContent = `Week ${week.week_num} Race Meetings (${week.events.length})`;
        subtitle.textContent = `${week.date_formatted} • ${week.status_desc}`;
        
        if (week.events.length >= 5) {
          tag.className = 'intensity-tag mega';
          tag.textContent = `🔥 Mega Super Weekend (${week.events.length} Series)`;
        } else if (week.events.length >= 3) {
          tag.className = 'intensity-tag multi';
          tag.textContent = `⚡ Multi-Series Weekend (${week.events.length} Series)`;
        } else {
          tag.className = 'intensity-tag normal';
          tag.textContent = `🏁 ${week.events.length} Championship Event${week.events.length > 1 ? 's' : ''}`;
        }

        grid.style.display = 'grid';
        offBanner.style.display = 'none';
        grid.innerHTML = '';

        week.events.forEach(ev => {
          const card = document.createElement('div');
          card.className = 'event-card';

          let badgesHtml = `<span class="tag">${ev.sub_series || ev.series}</span>`;
          if (ev.badge.includes('Sprint')) {
            badgesHtml += `<span class="tag tag-sprint">⚡ Sprint Race</span>`;
          }
          if (ev.badge.includes('Double Header')) {
            badgesHtml += `<span class="tag tag-double">✌️ Double Header</span>`;
          }
          if (ev.badge.includes('F3')) {
            badgesHtml += `<span class="tag tag-f3">🟢 F3 Support Round</span>`;
          }
          if (ev.badge.includes('F2')) {
            badgesHtml += `<span class="tag tag-f2">🔵 F2 Support Round</span>`;
          }
          if (ev.title.includes('24 Hours') || ev.title.includes('24h') || ev.title.includes('12 Hour') || ev.title.includes('Petit Le Mans') || ev.title.includes('1812km')) {
            badgesHtml += `<span class="tag tag-enduro">⏱️ Major Endurance Classic</span>`;
          }

          card.innerHTML = `
            <div class="card-top">
              <span class="series-badge ${getSeriesBadgeClass(ev.series)}">${ev.series}</span>
              <span class="refactored-inline-14">${ev.flag}</span>
            </div>
            <div class="card-title">${ev.title}</div>
            <div class="card-venue">
              <div class="circuit-name">📍 ${ev.circuit}</div>
              <div>${ev.city ? ev.city + ', ' : ''}${ev.country}</div>
            </div>
            <div class="badges-row">${badgesHtml}</div>
            ${ev.notes ? `<div class="refactored-inline-15">${ev.notes}</div>` : ''}
            <div class="card-footer">
              ${isSeriesConfirmed(ev.series) ? `<span class="refactored-inline-16">⏱️ Schedule: Confirmed</span>` : `<span class="refactored-inline-17">⏱️ Schedule: TBC</span>`}
              <span data-event="${encodeURIComponent(JSON.stringify(ev))}" class="refactored-inline-18">View Details ➔</span>
            </div>
          `;
          grid.appendChild(card);
        });
      }

      // Update adjacent cards
      const prevWeek = CALENDAR_DATA.find(w => w.week_num === currentWeekNum - 1);
      const nextWeek = CALENDAR_DATA.find(w => w.week_num === currentWeekNum + 1);

      if (prevWeek) {
        document.getElementById('cardPrevWeek').style.visibility = 'visible';
        document.getElementById('prevWeekTitle').textContent = `Week ${prevWeek.week_num} • ${prevWeek.date_formatted}`;
        document.getElementById('prevWeekCount').textContent = prevWeek.events.length > 0 ? `${prevWeek.events.length} Series` : 'Off Week';
      } else {
        document.getElementById('cardPrevWeek').style.visibility = 'hidden';
      }

      if (nextWeek) {
        document.getElementById('cardNextWeek').style.visibility = 'visible';
        document.getElementById('nextWeekTitle').textContent = `Week ${nextWeek.week_num} • ${nextWeek.date_formatted}`;
        document.getElementById('nextWeekCount').textContent = nextWeek.events.length > 0 ? `${nextWeek.events.length} Series` : 'Off Week';
      } else {
        document.getElementById('cardNextWeek').style.visibility = 'hidden';
      }
    }

    // Render Continuous Feed
    function renderFeed() {
      const container = document.getElementById('viewFeed');
      container.innerHTML = '';

      CALENDAR_DATA.forEach(week => {
        const card = document.createElement('div');
        card.className = 'week-row-card';

        let eventsHtml = '';
        if (week.events.length === 0) {
          eventsHtml = `<div class="refactored-inline-19">🌴 ${week.status_desc}</div>`;
        } else {
          week.events.forEach(ev => {
            eventsHtml += `
              <div class="event-chip" data-event="${encodeURIComponent(JSON.stringify(ev))}">
                <span class="series-badge ${getSeriesBadgeClass(ev.series)} refactored-inline-20">${ev.series}</span>
                <span>${ev.flag} <strong>${ev.raw_text}</strong></span>
                <span class="refactored-inline-21">${ev.circuit}</span>
              </div>
            `;
          });
        }

        card.innerHTML = `
          <div class="week-row-header">
            <div>
              <strong class="refactored-inline-22">Week ${week.week_num}</strong>
              <span class="refactored-inline-23">${week.date_formatted}</span>
            </div>
            <span class="tag ${week.events.length >= 5 ? 'tag-sprint' : (week.events.length >= 3 ? 'tag-double' : '')}">
              ${week.events.length > 0 ? `${week.events.length} Series` : 'Off Week'}
            </span>
          </div>
          <div class="week-row-events">${eventsHtml}</div>
        `;
        container.appendChild(card);
      });
    }

    // Render Master Excel Matrix Table
    function renderMatrix() {
      const table = document.getElementById('matrixTable');
      
      const headers = [
        'Week', 'Dates', 'F1 & F2 & F3 (Green-F3, Blue-F2)', 'FE', 'WEC / ELMS / ALMS', 
        'GTWCE / GTWCAs / IGTC', 'IMSA', 'DTM', 'WRC', 'FREC / FROT / FWS', 
        'Italian / Spanish / German F4'
      ];

      let theadHtml = '<thead><tr>';
      headers.forEach((h, i) => {
        if (i === 0) theadHtml += `<th class="sticky-col">${h}</th>`;
        else if (i === 1) theadHtml += `<th class="sticky-col-2">${h}</th>`;
        else theadHtml += `<th>${h}</th>`;
      });
      theadHtml += '</tr></thead>';

      let tbodyHtml = '<tbody>';
      CALENDAR_DATA.forEach(w => {
        tbodyHtml += `<tr>`;
        tbodyHtml += `<td class="sticky-col refactored-inline-24">W${w.week_num}</td>`;
        tbodyHtml += `<td class="sticky-col-2 refactored-inline-25">${w.start_date.substring(5)} to ${w.end_date.substring(5)}</td>`;

        // Col 3: F1
        const f1Ev = w.events.find(e => e.col === 3);
        tbodyHtml += renderMatrixCell(f1Ev);

        // Col 4: FE
        const feEv = w.events.find(e => e.col === 4);
        tbodyHtml += renderMatrixCell(feEv);

        // Col 5: WEC/ELMS/ALMS
        const wecEv = w.events.find(e => e.col === 5);
        tbodyHtml += renderMatrixCell(wecEv);

        // Col 6-9: GT
        const gtEvs = w.events.filter(e => 6 <= e.col && e.col <= 9);
        tbodyHtml += renderMatrixCellList(gtEvs);

        // Col 10: IMSA
        const imsaEv = w.events.find(e => e.col === 10);
        tbodyHtml += renderMatrixCell(imsaEv);

        // Col 11: DTM
        const dtmEv = w.events.find(e => e.col === 11);
        tbodyHtml += renderMatrixCell(dtmEv);

        // Col 12: WRC
        const wrcEv = w.events.find(e => e.col === 12);
        tbodyHtml += renderMatrixCell(wrcEv);

        // Col 13-14: FREC
        const frecEvs = w.events.filter(e => 13 <= e.col && e.col <= 14);
        tbodyHtml += renderMatrixCellList(frecEvs);

        // Col 15-18: F4
        const f4Evs = w.events.filter(e => 15 <= e.col && e.col <= 18);
        tbodyHtml += renderMatrixCellList(f4Evs);

        tbodyHtml += `</tr>`;
      });
      tbodyHtml += '</tbody>';

      table.innerHTML = theadHtml + tbodyHtml;
    }

    function renderMatrixCell(ev) {
      if (!ev) return '<td></td>';
      const cellClass = getCellClass(ev);
      return `<td class="${cellClass} refactored-inline-26" data-event="${encodeURIComponent(JSON.stringify(ev))}">${ev.flag} ${ev.raw_text}</td>`;
    }

    function renderMatrixCellList(evs) {
      if (!evs || evs.length === 0) return '<td></td>';
      const chips = evs.map(ev => {
        const cellClass = getCellClass(ev);
        return `<span class="${cellClass} refactored-inline-27" data-event="${encodeURIComponent(JSON.stringify(ev))}">${ev.flag} ${ev.raw_text}</span>`;
      }).join(' ');
      return `<td>${chips}</td>`;
    }

    // Render Championships View
    function renderChampionships(targetSeries) {
      selectedChampionship = targetSeries || selectedChampionship;
      
      const seriesList = [
        'Formula 1', 'Formula E', 'FIA WEC', 'ELMS', 'Asian Le Mans',
        'IGTC', 'GTWCE', 'GTWC Asia', 'IMSA', 'DTM', 'WRC',
        'FRECA', 'FROT', 'FWS', 'Italian F4', 'Spanish F4', 'German F4', 'Macau GP'
      ];

      const navContainer = document.getElementById('seriesNavPills');
      navContainer.innerHTML = '';
      seriesList.forEach(s => {
        const btn = document.createElement('button');
        btn.className = `series-pill-btn ${s === selectedChampionship ? 'active' : ''}`;
        btn.textContent = s;
        btn.onclick = () => renderChampionships(s);
        navContainer.appendChild(btn);
      });

      // Find all events for this series
      const scheduleContainer = document.getElementById('seriesScheduleList');
      scheduleContainer.innerHTML = '';

      let round = 1;
      CALENDAR_DATA.forEach(w => {
        w.events.filter(e => e.series === selectedChampionship).forEach(ev => {
          const item = document.createElement('div');
          item.className = 'series-round-item';
          item.innerHTML = `
            <div class="refactored-inline-28">
              <div class="round-num-badge">R${round++}</div>
              <div>
                <strong class="refactored-inline-29">${ev.title}</strong>
                <div class="refactored-inline-30">
                  📍 ${ev.circuit} • ${ev.city ? ev.city + ', ' : ''}${ev.country} ${ev.flag}
                </div>
              </div>
            </div>
            <div class="refactored-inline-31">
              <span class="tag">${w.date_formatted} (Week ${w.week_num})</span>
              ${ev.badge ? `<span class="tag tag-sprint">${ev.badge}</span>` : ''}
              <button class="btn refactored-inline-32" data-event="${encodeURIComponent(JSON.stringify(ev))}">Details</button>
            </div>
          `;
          scheduleContainer.appendChild(item);
        });
      });
    }

    // Render Super Weekends (Clashes)
    function renderClashes() {
      const container = document.getElementById('clashWeeksContainer');
      container.innerHTML = '';

      // Filter weeks with 4+ series
      const clashWeeks = CALENDAR_DATA.filter(w => w.events.length >= 4).sort((a, b) => b.events.length - a.events.length);

      clashWeeks.forEach(w => {
        const card = document.createElement('div');
        card.className = 'spotlight-card';
        card.innerHTML = `
          <div class="spotlight-header">
            <div>
              <h3>Week ${w.week_num} • ${w.date_formatted}</h3>
              <div class="spotlight-subtitle">${w.status_desc}</div>
            </div>
            <div class="intensity-tag ${w.events.length >= 6 ? 'mega' : 'multi'}">
              🔥 ${w.events.length} Simultaneous Championships
            </div>
          </div>
          <div class="events-grid">
            ${w.events.map(ev => `
              <div class="event-card refactored-inline-33" data-event="${encodeURIComponent(JSON.stringify(ev))}">
                <div class="card-top">
                  <span class="series-badge ${getSeriesBadgeClass(ev.series)}">${ev.series}</span>
                  <span>${ev.flag}</span>
                </div>
                <strong>${ev.title}</strong>
                <div class="refactored-inline-34">📍 ${ev.circuit}</div>
              </div>
            `).join('')}
          </div>
        `;
        container.appendChild(card);
      });
    }

    // Render Analytics
    function renderAnalytics() {
      // 1. Month Distribution
      const monthCounts = {};
      const monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
      monthNames.forEach(m => monthCounts[m] = 0);
      
      CALENDAR_DATA.forEach(w => {
        const m = monthNames[w.month_idx - 1];
        monthCounts[m] += w.events.length;
      });

      const maxMonth = Math.max(...Object.values(monthCounts));
      let monthHtml = '';
      Object.entries(monthCounts).forEach(([m, count]) => {
        const pct = (count / maxMonth) * 100;
        monthHtml += `
          <div class="bar-row">
            <span class="bar-label">${m} 2027</span>
            <div class="bar-track"><div class="bar-fill" style="width:${pct}%;"></div></div>
            <span class="bar-val">${count}</span>
          </div>
        `;
      });
      document.getElementById('chartMonthBars').innerHTML = monthHtml;

      // 2. Series Distribution
      const seriesCounts = {};
      CALENDAR_DATA.forEach(w => {
        w.events.forEach(e => {
          seriesCounts[e.series] = (seriesCounts[e.series] || 0) + 1;
        });
      });
      const sortedSeries = Object.entries(seriesCounts).sort((a, b) => b[1] - a[1]);
      const maxSeries = sortedSeries[0][1];
      let seriesHtml = '';
      sortedSeries.forEach(([s, count]) => {
        const pct = (count / maxSeries) * 100;
        seriesHtml += `
          <div class="bar-row">
            <span class="bar-label">${s}</span>
            <div class="bar-track"><div class="bar-fill" style="width:${pct}%; background:#0ea5e9;"></div></div>
            <span class="bar-val">${count}</span>
          </div>
        `;
      });
      document.getElementById('chartSeriesBars').innerHTML = seriesHtml;

      // 3. Top Circuits
      const venueCounts = {};
      CALENDAR_DATA.forEach(w => {
        w.events.forEach(e => {
          const v = e.circuit;
          venueCounts[v] = (venueCounts[v] || 0) + 1;
        });
      });
      const sortedVenues = Object.entries(venueCounts).sort((a, b) => b[1] - a[1]).slice(0, 8);
      const maxVenue = sortedVenues[0][1];
      let venueHtml = '';
      sortedVenues.forEach(([v, count]) => {
        const pct = (count / maxVenue) * 100;
        venueHtml += `
          <div class="bar-row">
            <span class="bar-label" title="${v}">${v}</span>
            <div class="bar-track"><div class="bar-fill" style="width:${pct}%; background:#10b981;"></div></div>
            <span class="bar-val">${count}</span>
          </div>
        `;
      });
      document.getElementById('chartVenueBars').innerHTML = venueHtml;

      // 4. Countries / Regions
      const countryCounts = {};
      CALENDAR_DATA.forEach(w => {
        w.events.forEach(e => {
          const c = e.country || 'International';
          countryCounts[c] = (countryCounts[c] || 0) + 1;
        });
      });
      const sortedCountries = Object.entries(countryCounts).sort((a, b) => b[1] - a[1]).slice(0, 8);
      const maxCountry = sortedCountries[0][1];
      let countryHtml = '';
      sortedCountries.forEach(([c, count]) => {
        const pct = (count / maxCountry) * 100;
        countryHtml += `
          <div class="bar-row">
            <span class="bar-label">${c}</span>
            <div class="bar-track"><div class="bar-fill" style="width:${pct}%; background:#8b5cf6;"></div></div>
            <span class="bar-val">${count}</span>
          </div>
        `;
      });
      document.getElementById('chartRegionBars').innerHTML = countryHtml;
    }

    // Modal popup
    
    const CONFIRMED_SERIES = new Set([
      'Formula 1', 'Formula E', 'FIA WEC', 'IMSA', 'DTM', 'WRC', 'IGTC', 'GTWCE', 'FROT', 'FWS'
    ]);

    function isSeriesConfirmed(series) {
      return CONFIRMED_SERIES.has(series);
    }

    function getEventTimetable(ev) {
      const s = ev.series || '';
      if (!isSeriesConfirmed(s)) {
        return null; // TBC: Schedule has not yet been announced
      }

      const raw = ev.raw_text || '';
      const title = ev.title || '';
      const badge = ev.badge || '';
      const isSprint = badge.includes('Sprint') || raw.includes('(s)');
      const hasF3 = badge.includes('F3') || (ev.sub_series && ev.sub_series.includes('F3'));
      const hasF2 = badge.includes('F2') || (ev.sub_series && ev.sub_series.includes('F2'));
      const isDoubleHeader = badge.includes('Double Header') || raw.includes('*2');

      // 1. Formula 1 (2027 Confirmed Schedule • 10 Sprints)
      if (s === 'Formula 1') {
        if (isSprint) {
          const days = [
            {
              day: 'Friday • Practice & Sprint Qualifying',
              sessions: [
                { time: '12:30 – 13:30', name: 'Formula 1 Free Practice 1 (FP1)', type: 'Practice' },
                { time: '16:30 – 17:14', name: 'Formula 1 Sprint Qualifying (SQ1, SQ2, SQ3)', type: 'Qualifying' }
              ]
            },
            {
              day: 'Saturday • Sprint Race & Grand Prix Qualifying',
              sessions: [
                { time: '12:00 – 12:45', name: 'Formula 1 Sprint Race (100 km / 19–24 Laps)', type: 'Sprint' },
                { time: '16:00 – 17:00', name: 'Formula 1 Grand Prix Qualifying (Q1, Q2, Q3)', type: 'Qualifying' }
              ]
            },
            {
              day: 'Sunday • Grand Prix Race Day',
              sessions: [
                { time: '13:00 – 13:30', name: 'Formula 1 Drivers\' Track Parade', type: 'Support' },
                { time: '15:00 – 17:00', name: `${title} (Main Race • 305 km)`, type: 'Race' }
              ]
            }
          ];

          if (hasF3 && hasF2) {
            days[0].sessions.unshift({ time: '09:55 – 10:40', name: 'FIA Formula 3 Free Practice', type: 'Support' });
            days[0].sessions.splice(1, 0, { time: '11:05 – 11:50', name: 'FIA Formula 2 Free Practice', type: 'Support' });
            days[0].sessions.push({ time: '15:00 – 15:30', name: 'FIA Formula 3 Qualifying Session', type: 'Support' });
            days[0].sessions.push({ time: '15:55 – 16:25', name: 'FIA Formula 2 Qualifying Session', type: 'Support' });

            days[1].sessions.unshift({ time: '10:15 – 11:00', name: 'FIA Formula 3 Sprint Race', type: 'Support' });
            days[1].sessions.splice(2, 0, { time: '14:15 – 15:05', name: 'FIA Formula 2 Sprint Race', type: 'Support' });

            days[2].sessions.unshift({ time: '08:30 – 09:15', name: 'FIA Formula 3 Feature Race', type: 'Support' });
            days[2].sessions.splice(1, 0, { time: '10:00 – 11:05', name: 'FIA Formula 2 Feature Race', type: 'Support' });
          } else if (hasF2) {
            days[0].sessions.unshift({ time: '11:05 – 11:50', name: 'FIA Formula 2 Free Practice', type: 'Support' });
            days[0].sessions.push({ time: '15:15 – 15:45', name: 'FIA Formula 2 Qualifying Session', type: 'Support' });
            days[1].sessions.splice(1, 0, { time: '14:15 – 15:05', name: 'FIA Formula 2 Sprint Race', type: 'Support' });
            days[2].sessions.unshift({ time: '10:00 – 11:05', name: 'FIA Formula 2 Feature Race', type: 'Support' });
          }
          return days;
        } else {
          // Standard F1 Weekend
          const isNightRace = title.includes('Singapore') || title.includes('Las Vegas') || title.includes('Saudi') || title.includes('Bahrain') || title.includes('Abu Dhabi');
          const raceTime = isNightRace ? '20:00 – 22:00' : '15:00 – 17:00';
          const fp2Time = isNightRace ? '19:00 – 20:00' : '17:00 – 18:00';
          const qualTime = isNightRace ? '19:00 – 20:00' : '16:00 – 17:00';

          const days = [
            {
              day: 'Friday • Free Practice 1 & 2',
              sessions: [
                { time: '13:30 – 14:30', name: 'Formula 1 Free Practice 1 (FP1)', type: 'Practice' },
                { time: fp2Time, name: 'Formula 1 Free Practice 2 (FP2)', type: 'Practice' }
              ]
            },
            {
              day: 'Saturday • Final Practice & Qualifying',
              sessions: [
                { time: '12:30 – 13:30', name: 'Formula 1 Free Practice 3 (FP3)', type: 'Practice' },
                { time: qualTime, name: 'Formula 1 Qualifying (Q1, Q2, Q3)', type: 'Qualifying' }
              ]
            },
            {
              day: 'Sunday • Grand Prix Race Day',
              sessions: [
                { time: '13:00 – 13:30', name: 'Formula 1 Drivers\' Track Parade', type: 'Support' },
                { time: raceTime, name: `${title} (Main Race • 305 km)`, type: 'Race' }
              ]
            }
          ];

          if (hasF3 && hasF2) {
            days[0].sessions.unshift({ time: '09:55 – 10:40', name: 'FIA Formula 3 Free Practice', type: 'Support' });
            days[0].sessions.splice(1, 0, { time: '11:05 – 11:50', name: 'FIA Formula 2 Free Practice', type: 'Support' });
            days[0].sessions.push({ time: '15:00 – 15:30', name: 'FIA Formula 3 Qualifying Session', type: 'Support' });
            days[0].sessions.push({ time: '15:55 – 16:25', name: 'FIA Formula 2 Qualifying Session', type: 'Support' });

            days[1].sessions.unshift({ time: '10:00 – 10:45', name: 'FIA Formula 3 Sprint Race', type: 'Support' });
            days[1].sessions.splice(2, 0, { time: '14:15 – 15:05', name: 'FIA Formula 2 Sprint Race', type: 'Support' });

            days[2].sessions.unshift({ time: '08:30 – 09:15', name: 'FIA Formula 3 Feature Race', type: 'Support' });
            days[2].sessions.splice(1, 0, { time: '10:00 – 11:05', name: 'FIA Formula 2 Feature Race', type: 'Support' });
          } else if (hasF2) {
            days[0].sessions.unshift({ time: '11:05 – 11:50', name: 'FIA Formula 2 Free Practice', type: 'Support' });
            days[0].sessions.push({ time: '15:15 – 15:45', name: 'FIA Formula 2 Qualifying Session', type: 'Support' });
            days[1].sessions.splice(1, 0, { time: '14:15 – 15:05', name: 'FIA Formula 2 Sprint Race', type: 'Support' });
            days[2].sessions.unshift({ time: '10:00 – 11:05', name: 'FIA Formula 2 Feature Race', type: 'Support' });
          }
          return days;
        }
      }

      // 2. Castrol Toyota Formula Regional Oceania Trophy (FROT - Confirmed!)
      if (s === 'FROT') {
        const isNZGP = title.includes('Highlands') || title.includes('NZ GP');
        return [
          {
            day: 'Friday • Free Practice & Testing',
            sessions: [
              { time: '10:30 – 11:15', name: 'Free Practice 1 (45 Mins)', type: 'Practice' },
              { time: '14:00 – 14:45', name: 'Free Practice 2 (45 Mins)', type: 'Practice' }
            ]
          },
          {
            day: 'Saturday • Qualifying 1 & Race 1',
            sessions: [
              { time: '10:15 – 10:30', name: 'Qualifying 1 (15 Mins • Sets Race 1 Grid)', type: 'Qualifying' },
              { time: '14:40 – 15:15', name: 'FROT Race 1 (18 Laps • Full Championship Points)', type: 'Race' }
            ]
          },
          {
            day: 'Sunday • Qualifying 2 & Double Races',
            sessions: [
              { time: '09:45 – 10:00', name: 'Qualifying 2 (15 Mins • Sets Feature Race Grid)', type: 'Qualifying' },
              { time: '11:30 – 12:05', name: 'FROT Race 2 (18 Laps • Reverse Top 8 Grid)', type: 'Sprint' },
              { time: '15:15 – 16:00', name: isNZGP ? '71st New Zealand Grand Prix (27 Laps)' : 'FROT Feature Race 3 (23 Laps)', type: 'Race' },
              { time: '16:15 – 16:45', name: 'Official Podium Ceremony & Trophy Presentation', type: 'Support' }
            ]
          }
        ];
      }

      // 3. Formula Winter Series (FWS - Confirmed!)
      if (s === 'FWS') {
        return [
          {
            day: 'Friday • Collective Testing',
            sessions: [
              { time: '09:30 – 10:30', name: 'Collective Test Session 1', type: 'Practice' },
              { time: '14:00 – 15:00', name: 'Collective Test Session 2', type: 'Practice' }
            ]
          },
          {
            day: 'Saturday • Qualifying 1 & Race 1',
            sessions: [
              { time: '09:50 – 10:10', name: 'Qualifying 1 (20 Mins)', type: 'Qualifying' },
              { time: '14:20 – 14:55', name: 'FWS Race 1 (30 Mins + 1 Lap)', type: 'Race' }
            ]
          },
          {
            day: 'Sunday • Qualifying 2 & Races 2 & 3',
            sessions: [
              { time: '09:30 – 09:50', name: 'Qualifying 2 (20 Mins)', type: 'Qualifying' },
              { time: '11:40 – 12:15', name: 'FWS Race 2 (30 Mins + 1 Lap)', type: 'Race' },
              { time: '15:20 – 15:55', name: 'FWS Race 3 (30 Mins + 1 Lap)', type: 'Race' },
              { time: '16:10', name: 'Podium Ceremony', type: 'Support' }
            ]
          }
        ];
      }

      // 4. Formula E (Season 13 Gen4 Confirmed!)
      if (s === 'Formula E') {
        if (isDoubleHeader) {
          return [
            {
              day: 'Friday • Shakedown & Practice 1',
              sessions: [
                { time: '14:30 – 15:00', name: 'Formula E Shakedown', type: 'Practice' },
                { time: '16:30 – 17:15', name: 'Free Practice 1 (FP1)', type: 'Practice' }
              ]
            },
            {
              day: 'Saturday • E-Prix Round 1',
              sessions: [
                { time: '07:30 – 08:15', name: 'Free Practice 2 (FP2)', type: 'Practice' },
                { time: '09:40 – 11:05', name: 'Round 1 Qualifying (Groups & Knockout Duels)', type: 'Qualifying' },
                { time: '14:03 – 15:15', name: 'Round 1 E-Prix Race (Attack Mode • 45m + 1 Lap)', type: 'Race' },
                { time: '15:30 – 16:00', name: 'Podium & Drivers\' Press Conference', type: 'Support' }
              ]
            },
            {
              day: 'Sunday • E-Prix Round 2',
              sessions: [
                { time: '07:30 – 08:15', name: 'Free Practice 3 (FP3)', type: 'Practice' },
                { time: '09:40 – 11:05', name: 'Round 2 Qualifying (Groups & Knockout Duels)', type: 'Qualifying' },
                { time: '14:03 – 15:15', name: 'Round 2 E-Prix Race (Attack Mode • 45m + 1 Lap)', type: 'Race' },
                { time: '15:30 – 16:00', name: 'Podium & Trophy Celebrations', type: 'Support' }
              ]
            }
          ];
        } else {
          return [
            {
              day: 'Friday • Shakedown & Practice',
              sessions: [
                { time: '16:30 – 17:15', name: 'Free Practice 1 (FP1)', type: 'Practice' }
              ]
            },
            {
              day: 'Saturday • E-Prix Race Day',
              sessions: [
                { time: '07:30 – 08:15', name: 'Free Practice 2 (FP2)', type: 'Practice' },
                { time: '09:40 – 11:05', name: 'Qualifying Session (Group Stages & Knockout Duels)', type: 'Qualifying' },
                { time: '14:03 – 15:15', name: `${title} (Main Race • 45m + 1 Lap)`, type: 'Race' },
                { time: '15:30 – 16:00', name: 'Podium Ceremony', type: 'Support' }
              ]
            }
          ];
        }
      }

      // 5. FIA WEC (Confirmed!)
      if (s === 'FIA WEC') {
        if (title.includes('Le Mans')) {
          return [
            {
              day: 'Wednesday • Qualifying Practice & Night Session',
              sessions: [
                { time: '14:00 – 17:00', name: 'Free Practice 1 (FP1)', type: 'Practice' },
                { time: '19:00 – 20:00', name: 'Qualifying Practice (Top 8 Hyperpole Qualifiers)', type: 'Qualifying' },
                { time: '22:00 – 00:00', name: 'Free Practice 2 (Night Practice)', type: 'Practice' }
              ]
            },
            {
              day: 'Thursday • Hyperpole Shootout',
              sessions: [
                { time: '15:00 – 18:00', name: 'Free Practice 3 (FP3)', type: 'Practice' },
                { time: '20:00 – 20:30', name: 'Hyperpole Shootout (Hypercar & LMGT3 Battle)', type: 'Qualifying' },
                { time: '22:00 – 23:00', name: 'Free Practice 4 (Final Night Practice)', type: 'Practice' }
              ]
            },
            {
              day: 'Friday • Drivers\' Parade in Le Mans Town',
              sessions: [
                { time: '14:00 – 16:30', name: 'Classic Drivers\' Parade (Downtown Le Mans)', type: 'Support' }
              ]
            },
            {
              day: 'Saturday • The 24-Hour Start',
              sessions: [
                { time: '09:30 – 09:45', name: 'Warm-Up Session', type: 'Practice' },
                { time: '15:30 – 16:00', name: 'Starting Grid Ceremony & French Tricolour Drop', type: 'Support' },
                { time: '16:00', name: '95th 24 Hours of Le Mans Official Race Start', type: 'Race' }
              ]
            },
            {
              day: 'Sunday • 24-Hour Finish',
              sessions: [
                { time: '16:00', name: '24 Hours of Le Mans Chequered Flag & Finish', type: 'Race' },
                { time: '16:15 – 17:00', name: 'Historic Winners\' Podium Ceremony', type: 'Support' }
              ]
            }
          ];
        } else {
          return [
            {
              day: 'Friday • Free Practices',
              sessions: [
                { time: '11:00 – 12:30', name: 'Free Practice 1 (90 Mins)', type: 'Practice' },
                { time: '16:00 – 17:30', name: 'Free Practice 2 (90 Mins)', type: 'Practice' }
              ]
            },
            {
              day: 'Saturday • Practice 3 & Hyperpole Qualifying',
              sessions: [
                { time: '10:45 – 11:45', name: 'Free Practice 3 (60 Mins)', type: 'Practice' },
                { time: '14:40 – 15:05', name: 'LMGT3 Qualifying & Hyperpole', type: 'Qualifying' },
                { time: '15:20 – 15:45', name: 'Hypercar Qualifying & Hyperpole', type: 'Qualifying' }
              ]
            },
            {
              day: 'Sunday • Endurance Race Day',
              sessions: [
                { time: '11:00 / 12:00', name: `${title} (Official Race Start)`, type: 'Race' },
                { time: 'Post-Race', name: 'Official Podium Ceremony & Trophy Presentation', type: 'Support' }
              ]
            }
          ];
        }
      }

      // 6. IMSA WeatherTech SportsCar Championship (Confirmed!)
      if (s === 'IMSA') {
        if (title.includes('Rolex 24')) {
          return [
            {
              day: 'Thursday • Practice & Night Practice',
              sessions: [
                { time: '10:05 – 11:35', name: 'Practice 1', type: 'Practice' },
                { time: '14:10 – 15:55', name: 'Practice 2 & Qualifying', type: 'Qualifying' },
                { time: '18:35 – 20:05', name: 'Night Practice (Full Field)', type: 'Practice' }
              ]
            },
            {
              day: 'Friday • Final Shakedown',
              sessions: [
                { time: '11:20 – 12:20', name: 'Practice 4 (Final Shakedown)', type: 'Practice' }
              ]
            },
            {
              day: 'Saturday • Rolex 24 Start',
              sessions: [
                { time: '11:30 – 13:00', name: 'Fan Grid Walk & Pre-Race Festivities', type: 'Support' },
                { time: '13:40', name: '65th Rolex 24 at Daytona Official Race Start', type: 'Race' }
              ]
            },
            {
              day: 'Sunday • Rolex 24 Finish',
              sessions: [
                { time: '13:40', name: 'Rolex 24 at Daytona Chequered Flag & Finish', type: 'Race' },
                { time: '14:00', name: 'Gatorade Victory Lane Celebration', type: 'Support' }
              ]
            }
          ];
        } else if (title.includes('Sebring')) {
          return [
            {
              day: 'Thursday • Practice & Night Practice',
              sessions: [
                { time: '10:05 – 11:35', name: 'Practice 1', type: 'Practice' },
                { time: '15:05 – 16:35', name: 'Practice 2', type: 'Practice' },
                { time: '19:45 – 21:15', name: 'Night Practice (12h Simulation)', type: 'Practice' }
              ]
            },
            {
              day: 'Friday • Qualifying',
              sessions: [
                { time: '09:00 – 10:15', name: 'Qualifying (GTP, LMP2, GTD PRO, GTD)', type: 'Qualifying' }
              ]
            },
            {
              day: 'Saturday • 12 Hours of Sebring',
              sessions: [
                { time: '08:00 – 08:20', name: 'Warm-Up Session', type: 'Practice' },
                { time: '09:40 – 21:40', name: '75th Mobil 1 Twelve Hours of Sebring (Race Start to Finish)', type: 'Race' },
                { time: '21:50', name: 'Victory Lane Fireworks & Celebrations', type: 'Support' }
              ]
            }
          ];
        } else {
          return [
            {
              day: 'Friday • Practice 1',
              sessions: [
                { time: '14:00 – 15:30', name: 'Practice 1 (90 Mins)', type: 'Practice' }
              ]
            },
            {
              day: 'Saturday • Practice 2 & Qualifying',
              sessions: [
                { time: '09:30 – 11:00', name: 'Practice 2', type: 'Practice' },
                { time: '13:15 – 14:30', name: 'Qualifying (All Classes)', type: 'Qualifying' }
              ]
            },
            {
              day: 'Sunday • Main Race Day',
              sessions: [
                { time: '08:30 – 08:50', name: 'Warm-Up Session', type: 'Practice' },
                { time: '12:00 – 13:00', name: 'Open Grid & Fan Walk', type: 'Support' },
                { time: '13:10 – 15:50', name: `${title} (Main Race • 2h 40m)`, type: 'Race' },
                { time: '16:00', name: 'Victory Lane Ceremony', type: 'Support' }
              ]
            }
          ];
        }
      }

      // 7. DTM (Confirmed!)
      if (s === 'DTM') {
        return [
          {
            day: 'Friday • Free Practices',
            sessions: [
              { time: '11:00 – 11:45', name: 'Free Practice 1 (45 Mins)', type: 'Practice' },
              { time: '14:45 – 15:30', name: 'Free Practice 2 (45 Mins)', type: 'Practice' }
            ]
          },
          {
            day: 'Saturday • Qualifying 1 & Race 1',
            sessions: [
              { time: '09:30 – 09:50', name: 'Qualifying 1 (20 Mins)', type: 'Qualifying' },
              { time: '13:30 – 14:30', name: 'DTM Race 1 (60 Mins + Mandatory Pit Stop)', type: 'Race' },
              { time: '14:45', name: 'Podium & Press Conference', type: 'Support' }
            ]
          },
          {
            day: 'Sunday • Qualifying 2 & Race 2',
            sessions: [
              { time: '09:30 – 09:50', name: 'Qualifying 2 (20 Mins)', type: 'Qualifying' },
              { time: '13:30 – 14:30', name: 'DTM Race 2 (60 Mins + Mandatory Pit Stop)', type: 'Race' },
              { time: '14:45', name: 'Podium & Trophy Ceremony', type: 'Support' }
            ]
          }
        ];
      }

      // 8. WRC (Confirmed!)
      if (s === 'WRC') {
        return [
          {
            day: 'Thursday • Shakedown & Super Special 1',
            sessions: [
              { time: '09:01 – 13:30', name: 'Official Shakedown Stage', type: 'Practice' },
              { time: '18:05 – 20:00', name: 'Ceremonial Start & Super Special Stage 1 (SSS1)', type: 'Sprint' }
            ]
          },
          {
            day: 'Friday • Leg 1 (Forest / Mountain Stages)',
            sessions: [
              { time: '07:30 – 12:30', name: 'Morning Loop (Special Stages SS2 – SS4)', type: 'Race' },
              { time: '12:30 – 13:15', name: 'Midday Flexi-Service Park (40 Mins)', type: 'Support' },
              { time: '13:45 – 18:30', name: 'Afternoon Loop (Special Stages SS5 – SS7)', type: 'Race' }
            ]
          },
          {
            day: 'Saturday • Leg 2 (Longest Loop)',
            sessions: [
              { time: '07:15 – 12:00', name: 'Morning Loop (Special Stages SS8 – SS11)', type: 'Race' },
              { time: '12:00 – 12:45', name: 'Midday Flexi-Service Park (40 Mins)', type: 'Support' },
              { time: '13:15 – 18:45', name: 'Afternoon Loop (Special Stages SS12 – SS15)', type: 'Race' }
            ]
          },
          {
            day: 'Sunday • Super Sunday & Wolf Power Stage',
            sessions: [
              { time: '07:30 – 11:30', name: 'Morning Stages (Special Stages SS16 – SS18)', type: 'Race' },
              { time: '12:15 – 13:30', name: 'Wolf Power Stage (SS19 • Live TV Finale with Bonus Points)', type: 'Sprint' },
              { time: '14:15 – 15:00', name: 'Official WRC Podium & Trophy Ceremony', type: 'Support' }
            ]
          }
        ];
      }

      // 9. IGTC & GTWCE (SRO Confirmed!)
      if (s === 'IGTC' || s === 'GTWCE') {
        if (title.includes('24 Hours of Spa') || title.includes('24H Nuerburgring')) {
          return [
            {
              day: 'Thursday • Practice & Night Qualifying',
              sessions: [
                { time: '11:20 – 12:50', name: 'Free Practice (90 Mins)', type: 'Practice' },
                { time: '16:10 – 17:10', name: 'Pre-Qualifying Session', type: 'Practice' },
                { time: '20:55 – 22:15', name: 'Qualifying Session (Q1–Q4)', type: 'Qualifying' },
                { time: '22:55 – 00:25', name: 'Night Practice (90 Mins)', type: 'Practice' }
              ]
            },
            {
              day: 'Friday • Super Pole Shootout',
              sessions: [
                { time: '15:35 – 16:05', name: 'Warm-Up Session', type: 'Practice' },
                { time: '16:30 – 17:15', name: 'Super Pole Shootout (Top 20 GT3 Fast Laps)', type: 'Qualifying' }
              ]
            },
            {
              day: 'Saturday • Race Start',
              sessions: [
                { time: '15:30 – 16:30', name: 'Starting Grid Ceremony & Pit Walk', type: 'Support' },
                { time: '16:30', name: `${title} Official 24-Hour Start`, type: 'Race' }
              ]
            },
            {
              day: 'Sunday • 24-Hour Finish',
              sessions: [
                { time: '16:30', name: `${title} Chequered Flag & Finish`, type: 'Race' },
                { time: '16:45 – 17:30', name: 'Victory Lane Celebration', type: 'Support' }
              ]
            }
          ];
        } else if (isSprint) {
          return [
            {
              day: 'Friday • Free Practice & Pre-Qualifying',
              sessions: [
                { time: '09:00 – 10:20', name: 'Free Practice (80 Mins)', type: 'Practice' },
                { time: '14:10 – 15:30', name: 'Pre-Qualifying (80 Mins)', type: 'Practice' }
              ]
            },
            {
              day: 'Saturday • Qualifying 1 & Race 1',
              sessions: [
                { time: '09:00 – 09:35', name: 'Qualifying 1 (Grid for Race 1)', type: 'Qualifying' },
                { time: '14:00 – 15:00', name: 'Sprint Race 1 (60 Mins with Driver Change)', type: 'Race' }
              ]
            },
            {
              day: 'Sunday • Qualifying 2 & Race 2',
              sessions: [
                { time: '09:00 – 09:35', name: 'Qualifying 2 (Grid for Race 2)', type: 'Qualifying' },
                { time: '14:00 – 15:00', name: 'Sprint Race 2 (60 Mins with Driver Change)', type: 'Race' },
                { time: '15:15 – 15:45', name: 'Podium Ceremony', type: 'Support' }
              ]
            }
          ];
        } else {
          return [
            {
              day: 'Friday • Bronze Test & Practice',
              sessions: [
                { time: '10:00 – 11:00', name: 'Bronze Driver Test Session', type: 'Practice' },
                { time: '14:30 – 16:00', name: 'Free Practice (90 Mins)', type: 'Practice' }
              ]
            },
            {
              day: 'Saturday • Pre-Qualifying',
              sessions: [
                { time: '09:00 – 10:30', name: 'Pre-Qualifying Session (90 Mins)', type: 'Practice' }
              ]
            },
            {
              day: 'Sunday • Qualifying & Endurance Race',
              sessions: [
                { time: '09:45 – 10:45', name: 'Qualifying 1, 2 & 3 (Combined Times)', type: 'Qualifying' },
                { time: '15:00 – 18:00', name: `${title} (Endurance Race)`, type: 'Race' },
                { time: '18:15', name: 'Podium Ceremony', type: 'Support' }
              ]
            }
          ];
        }
      }

      return null;
    }
    window.showEventModal = function(ev) {
      const modal = document.getElementById('eventModal');
      const content = document.getElementById('modalContent');
      
      const timetable = getEventTimetable(ev);
      const isConfirmed = timetable !== null;

      let timetableSectionHtml = '';

      if (isConfirmed) {
        let timetableHtml = '';
        let totalSessions = 0;
        timetable.forEach(day => {
          totalSessions += day.sessions.length;
          let sessionsHtml = '';
          day.sessions.forEach(sess => {
            let badgeClass = 'badge-practice';
            if (sess.type === 'Qualifying') badgeClass = 'badge-qualifying';
            else if (sess.type === 'Sprint') badgeClass = 'badge-sprint';
            else if (sess.type === 'Race') badgeClass = 'badge-race';
            else if (sess.type === 'Support') badgeClass = 'badge-support';

            sessionsHtml += `
              <div class="timetable-session-row">
                <span class="session-time">${sess.time}</span>
                <span class="session-name" title="${sess.name}">${sess.name}</span>
                <span class="session-type-badge ${badgeClass}">${sess.type}</span>
              </div>
            `;
          });

          timetableHtml += `
            <div class="timetable-day-card">
              <div class="timetable-day-title">
                <span>📅 ${day.day}</span>
                <span class="refactored-inline-35">${day.sessions.length} Sessions</span>
              </div>
              <div class="timetable-session-list">
                ${sessionsHtml}
              </div>
            </div>
          `;
        });

        const scrollHint = timetable.length >= 3 
          ? `<span class="refactored-inline-36">📜 Scroll to view all ${timetable.length} days (${totalSessions} sessions)</span>`
          : `<span class="refactored-inline-37">${totalSessions} Total Sessions</span>`;

        timetableSectionHtml = `
          <div class="timetable-section">
            <div class="timetable-header">
              <div class="refactored-inline-38">
                <h4>⏱️ Official Weekend Timetable</h4>
                <span class="timetable-tz-notice refactored-inline-39">✓ Confirmed</span>
              </div>
              ${scrollHint}
            </div>
            <div class="timetable-days">
              ${timetableHtml}
            </div>
          </div>
        `;
      } else {
        timetableSectionHtml = `
          <div class="timetable-section">
            <div class="timetable-header">
              <h4>⏱️ Weekend Timetable</h4>
              <span class="timetable-tz-notice refactored-inline-40">Schedule: TBC</span>
            </div>
            <div class="tbc-box">
              <div class="tbc-title">TBC</div>
              <div class="tbc-desc">
                Official 2027 calendar dates and session timetables for <strong>${ev.series}</strong> have not yet been finalized by the sanctioning body. The official schedule will be updated once formally confirmed.
              </div>
            </div>
          </div>
        `;
      }

      content.innerHTML = `
        <div class="modal-header-area">
          <div class="modal-header-top">
            <span class="series-badge ${getSeriesBadgeClass(ev.series)}">${ev.series}</span>
            <span class="refactored-inline-41">${ev.flag}</span>
          </div>
          <div class="modal-title">${ev.title}</div>
          <div class="modal-venue-line">
            📍 <strong>${ev.circuit}</strong> • ${ev.city ? ev.city + ', ' : ''}${ev.country}
          </div>
        </div>
        
        <div class="event-meta-box">
          <div><strong>Championship:</strong> ${ev.sub_series || ev.series}</div>
          <div><strong>Format:</strong> ${ev.badge || 'Standard Championship Round'}</div>
          ${ev.notes ? `<div><strong>Notes:</strong> ${ev.notes}</div>` : ''}
        </div>

        ${timetableSectionHtml}

        <div class="modal-footer">
          <button class="btn btn-primary refactored-inline-42" data-close-modal="true">Close Window</button>
        </div>
      `;

      // Reset scroll position to top when opened
      const modalCard = document.querySelector('.modal-card');
      if (modalCard) modalCard.scrollTop = 0;

      modal.style.display = 'flex';
    };
// Export .ics calendar generator
    function exportICS() {
      let icsContent = "BEGIN:VCALENDAR\r\nVERSION:2.0\r\nPRODID:-//2027 Motorsports Calendar//EN\r\nCALSCALE:GREGORIAN\r\n";
      
      CALENDAR_DATA.forEach(w => {
        w.events.forEach(ev => {
          const startDt = w.start_date.replace(/-/g, '');
          // End date in ics is exclusive, add 1 day
          const endObj = new Date(w.end_date);
          endObj.setDate(endObj.getDate() + 1);
          const endDt = endObj.toISOString().slice(0, 10).replace(/-/g, '');

          icsContent += "BEGIN:VEVENT\r\n";
          icsContent += `UID:${w.week_num}-${ev.col}-2027@motorsportscalendar\r\n`;
          icsContent += `DTSTART;VALUE=DATE:${startDt}\r\n`;
          icsContent += `DTEND;VALUE=DATE:${endDt}\r\n`;
          icsContent += `SUMMARY:[${ev.series}] ${ev.title}\r\n`;
          icsContent += `LOCATION:${ev.circuit}\, ${ev.country}\r\n`;
          icsContent += `DESCRIPTION:${ev.title} (${ev.series} - ${ev.sub_series}). Circuit: ${ev.circuit}\, ${ev.city}\, ${ev.country}. Week ${w.week_num}.\r\n`;
          icsContent += "STATUS:CONFIRMED\r\n";
          icsContent += "END:VEVENT\r\n";
        });
      });
      
      icsContent += "END:VCALENDAR\r\n";

      const blob = new Blob([icsContent], { type: 'text/calendar;charset=utf-8;' });
      const link = document.createElement('a');
      link.href = URL.createObjectURL(blob);
      link.setAttribute('download', '2027_Motorsports_Calendar.ics');
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
    }

    // Export CSV generator
    function exportCSV() {
      let csv = "Week,Start Date,End Date,Series,Event Title,Circuit,City,Country,Notes,Excel Cell\n";
      CALENDAR_DATA.forEach(w => {
        w.events.forEach(ev => {
          const title = `"${ev.title.replace(/"/g, '""')}"`;
          const circuit = `"${ev.circuit.replace(/"/g, '""')}"`;
          const notes = `"${(ev.notes || '').replace(/"/g, '""')}"`;
          csv += `${w.week_num},${w.start_date},${w.end_date},${ev.series},${title},${circuit},"${ev.city}","${ev.country}",${notes},R${ev.row}C${ev.col}\n`;
        });
      });

      const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
      const link = document.createElement('a');
      link.href = URL.createObjectURL(blob);
      link.setAttribute('download', '2027_Motorsports_Calendar.csv');
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
    }

    // DOM Event Listeners & Bootstrapping
    document.addEventListener('DOMContentLoaded', () => {
      // Scrubber slider
      const scrubber = document.getElementById('weekScrubber');
      scrubber.addEventListener('input', (e) => {
        renderSpotlight(parseInt(e.target.value));
      });

      // Next / Prev buttons
      document.getElementById('btnPrevWeek').onclick = () => renderSpotlight(currentWeekNum - 1);
      document.getElementById('btnNextWeek').onclick = () => renderSpotlight(currentWeekNum + 1);
      document.getElementById('cardPrevWeek').onclick = () => renderSpotlight(currentWeekNum - 1);
      document.getElementById('cardNextWeek').onclick = () => renderSpotlight(currentWeekNum + 1);

      // Keyboard arrow shortcuts
      window.addEventListener('keydown', (e) => {
        if (e.target.tagName === 'INPUT' || e.target.tagName === 'SELECT') return;
        if (e.key === 'ArrowLeft') {
          renderSpotlight(currentWeekNum - 1);
        } else if (e.key === 'ArrowRight') {
          renderSpotlight(currentWeekNum + 1);
        }
      });

      // Quick jump pills
      document.querySelectorAll('.pill-btn').forEach(btn => {
        btn.onclick = () => {
          renderSpotlight(parseInt(btn.dataset.week));
        };
      });

      // View Switching Tabs
      const tabs = document.querySelectorAll('.tab-btn');
      tabs.forEach(tab => {
        tab.onclick = () => {
          tabs.forEach(t => t.classList.remove('active'));
          tab.classList.add('active');
          const view = tab.dataset.view;
          activeView = view;

          document.getElementById('viewSpotlight').style.display = view === 'spotlight' ? 'flex' : 'none';
          document.getElementById('viewFeed').style.display = view === 'feed' ? 'flex' : 'none';
          document.getElementById('viewMatrix').style.display = view === 'matrix' ? 'flex' : 'none';
          document.getElementById('viewChampionships').style.display = view === 'championships' ? 'flex' : 'none';
          document.getElementById('viewClashes').style.display = view === 'clashes' ? 'flex' : 'none';
          document.getElementById('viewAnalytics').style.display = view === 'analytics' ? 'grid' : 'none';

          if (view === 'feed') renderFeed();
          if (view === 'matrix') { renderMatrix(); applyFilters(); }
          if (view === 'championships') renderChampionships();
          if (view === 'clashes') renderClashes();
          if (view === 'analytics') renderAnalytics();
        };
      });

      // Light/Dark Theme Toggle
      const themeBtn = document.getElementById('btnThemeToggle');
      function updateThemeBtn(isLight) {
        if (isLight) {
          themeBtn.innerHTML = '🌙 Dark Mode';
          themeBtn.title = 'Switch to Dark Mode';
        } else {
          themeBtn.innerHTML = '☀️ Light Mode';
          themeBtn.title = 'Switch to Light Mode';
        }
      }

      // Check saved preference
      try {
        const savedTheme = localStorage.getItem('motorsports_2027_theme');
        if (savedTheme === 'light') {
          document.body.classList.add('light-mode');
          updateThemeBtn(true);
        } else {
          updateThemeBtn(false);
        }
      } catch (e) {
        updateThemeBtn(false);
      }

      themeBtn.onclick = () => {
        const isLight = document.body.classList.toggle('light-mode');
        try {
          localStorage.setItem('motorsports_2027_theme', isLight ? 'light' : 'dark');
        } catch (e) {}
        updateThemeBtn(isLight);
      };

      // Exports
      document.getElementById('btnExportICS').onclick = exportICS;
      document.getElementById('btnExportCSV').onclick = exportCSV;

      // Jump to next upcoming race
      document.getElementById('btnJumpNext').onclick = () => {
        const nextRaceWeek = CALENDAR_DATA.find(w => w.week_num > currentWeekNum && w.events.length > 0) || CALENDAR_DATA.find(w => w.events.length > 0);
        if (nextRaceWeek) {
          renderSpotlight(nextRaceWeek.week_num);
        }
      };

      // Filters
      const searchInput = document.getElementById('globalSearch');
      const seriesSelect = document.getElementById('seriesFilter');
      const monthSelect = document.getElementById('monthFilter');
      const specialSelect = document.getElementById('specialFilter');

      function applyFilters() {
        const query = searchInput.value.toLowerCase().trim();
        const selectedSeries = seriesSelect.value;
        const selectedMonth = monthSelect.value;
        const selectedSpecial = specialSelect.value;

        // 1. If in spotlight view, jump to first matching week
        if (activeView === 'spotlight') {
          const matchingWeek = CALENDAR_DATA.find(w => {
            if (selectedMonth !== 'ALL' && w.month_idx.toString() !== selectedMonth) return false;
            if (w.events.length === 0) return query === '' && selectedSeries === 'ALL' && selectedSpecial === 'ALL';
            
            return w.events.some(e => {
              if (selectedSeries !== 'ALL' && e.series !== selectedSeries) return false;
              if (selectedSpecial === 'SPRINT' && !e.badge.includes('Sprint')) return false;
              if (selectedSpecial === 'DOUBLE_HEADER' && !e.badge.includes('Double Header')) return false;
              if (selectedSpecial === 'SUPPORT_F3' && !e.badge.includes('F3')) return false;
              if (selectedSpecial === 'SUPPORT_F2' && !e.badge.includes('F2')) return false;
              if (selectedSpecial === 'MEGA' && w.events.length < 4) return false;
              if (selectedSpecial === 'ENDURANCE_24H' && !(e.title.includes('24 Hours') || e.title.includes('24h') || e.title.includes('12 Hour') || e.title.includes('Petit Le Mans'))) return false;

              if (query) {
                const matchStr = `${e.title} ${e.circuit} ${e.city} ${e.country} ${e.series} ${e.raw_text}`.toLowerCase();
                return matchStr.includes(query);
              }
              return true;
            });
          });

          if (matchingWeek) {
            renderSpotlight(matchingWeek.week_num);
          }
        } 
        
        // 2. If in feed view, filter rows
        else if (activeView === 'feed') {
          const rows = document.querySelectorAll('.week-row-card');
          rows.forEach((row, idx) => {
            const w = CALENDAR_DATA[idx];
            let show = true;
            if (selectedMonth !== 'ALL' && w.month_idx.toString() !== selectedMonth) show = false;
            if (show && selectedSpecial === 'MEGA' && w.events.length < 4) show = false;
            
            if (show && (query || selectedSeries !== 'ALL' || selectedSpecial !== 'ALL')) {
              const hasMatch = w.events.some(e => {
                if (selectedSeries !== 'ALL' && e.series !== selectedSeries) return false;
                if (selectedSpecial === 'SPRINT' && !e.badge.includes('Sprint')) return false;
                if (selectedSpecial === 'DOUBLE_HEADER' && !e.badge.includes('Double Header')) return false;
                if (selectedSpecial === 'SUPPORT_F3' && !e.badge.includes('F3')) return false;
                if (selectedSpecial === 'SUPPORT_F2' && !e.badge.includes('F2')) return false;
                if (selectedSpecial === 'ENDURANCE_24H' && !(e.title.includes('24 Hours') || e.title.includes('24h') || e.title.includes('12 Hour'))) return false;
                if (query) {
                  const matchStr = `${e.title} ${e.circuit} ${e.city} ${e.country} ${e.series} ${e.raw_text}`.toLowerCase();
                  return matchStr.includes(query);
                }
                return true;
              });
              show = hasMatch;
            }
            row.style.display = show ? 'flex' : 'none';
          });
        }

        // 3. If in matrix view, filter matrix table rows and highlight matching cells
        else if (activeView === 'matrix') {
          const rows = document.querySelectorAll('#matrixTable tbody tr');
          rows.forEach((row, idx) => {
            const w = CALENDAR_DATA[idx];
            let show = true;
            if (selectedMonth !== 'ALL' && w.month_idx.toString() !== selectedMonth) show = false;
            if (show && selectedSpecial === 'MEGA' && w.events.length < 4) show = false;

            if (show && (query || selectedSeries !== 'ALL' || selectedSpecial !== 'ALL')) {
              const hasMatch = w.events.some(e => {
                if (selectedSeries !== 'ALL' && e.series !== selectedSeries) return false;
                if (selectedSpecial === 'SPRINT' && !e.badge.includes('Sprint')) return false;
                if (selectedSpecial === 'DOUBLE_HEADER' && !e.badge.includes('Double Header')) return false;
                if (selectedSpecial === 'SUPPORT_F3' && !e.badge.includes('F3')) return false;
                if (selectedSpecial === 'SUPPORT_F2' && !e.badge.includes('F2')) return false;
                if (selectedSpecial === 'ENDURANCE_24H' && !(e.title.includes('24 Hours') || e.title.includes('24h') || e.title.includes('12 Hour'))) return false;
                if (query) {
                  const matchStr = `${e.title} ${e.circuit} ${e.city} ${e.country} ${e.series} ${e.raw_text}`.toLowerCase();
                  return matchStr.includes(query);
                }
                return true;
              });
              show = hasMatch;
            }
            row.style.display = show ? '' : 'none';
          });
        }
      }

      searchInput.addEventListener('input', applyFilters);
      seriesSelect.addEventListener('change', applyFilters);
      monthSelect.addEventListener('change', applyFilters);
      specialSelect.addEventListener('change', applyFilters);

      // Delegate event details and modal close actions for generated content.
      document.addEventListener('click', (e) => {
        const eventTrigger = e.target.closest('[data-event]');
        if (eventTrigger) {
          try { showEventModal(JSON.parse(decodeURIComponent(eventTrigger.dataset.event))); }
          catch (error) { console.error('Unable to open calendar event details.', error); }
          return;
        }
        if (e.target.closest('[data-close-modal]')) {
          document.getElementById('eventModal').style.display = 'none';
        }
      });

      // Modal close handlers
      document.getElementById('modalCloseBtn').onclick = () => {
        document.getElementById('eventModal').style.display = 'none';
      };
      window.onclick = (e) => {
        if (e.target === document.getElementById('eventModal')) {
          document.getElementById('eventModal').style.display = 'none';
        }
      };

      // Initial Render
      renderSpotlight(18); // Default to Week 18 (Miami Grand Prix & Mega Weekend)
    });
