"""Step 15e: Counties Reference — final 11 entries (South Dakota through DC)."""
import openpyxl
from openpyxl.styles import Font, Border, Side

def thin():
    s = Side(style="thin", color="CCCCCC")
    return Border(left=s, right=s, top=s, bottom=s)

counties = {
    'South Dakota': 'Aurora,Beadle,Bennett,Bon Homme,Brookings,Brown,Brule,Buffalo,Butte,Campbell,Charles Mix,Clark,Clay,Codington,Corson,Custer,Davison,Day,Deuel,Dewey,Douglas,Edmunds,Fall River,Faulk,Grant,Gregory,Haakon,Hamlin,Hand,Hanson,Harding,Hughes,Hutchinson,Hyde,Jackson,Jerauld,Jones,Kingsbury,Lake,Lawrence,Lincoln,Lyman,Marshall,McCook,McPherson,Meade,Mellette,Miner,Minnehaha,Moody,Pennington,Perkins,Potter,Roberts,Sanborn,Spink,Stanley,Sully,Todd,Tripp,Turner,Union,Walworth,Yankton,Ziebach',
    'Tennessee': 'Anderson,Bedford,Benton,Bledsoe,Blount,Bradley,Campbell,Cannon,Carroll,Carter,Cheatham,Chester,Claiborne,Clay,Cocke,Coffee,Crockett,Cumberland,Davidson,Decatur,DeKalb,Dickson,Dyer,Fayette,Fentress,Franklin,Gibson,Giles,Grainger,Greene,Grundy,Hamblen,Hamilton,Hancock,Hardeman,Hardin,Hawkins,Haywood,Henderson,Henry,Hickman,Houston,Humphreys,Jackson,Jefferson,Johnson,Knox,Lake,Lauderdale,Lawrence,Lewis,Lincoln,Loudon,Macon,Madison,Marion,Marshall,Maury,McMinn,McNairy,Meigs,Monroe,Montgomery,Moore,Morgan,Obion,Overton,Perry,Pickett,Polk,Putnam,Rhea,Roane,Robertson,Rutherford,Scott,Sequatchie,Sevier,Shelby,Smith,Stewart,Sullivan,Sumner,Tipton,Trousdale,Unicoi,Union,Van Buren,Warren,Washington,Wayne,Weakley,White,Williamson,Wilson',
    'Texas': 'Anderson,Andrews,Angelina,Aransas,Archer,Armstrong,Atascosa,Austin,Bailey,Bandera,Bastrop,Baylor,Bee,Bell,Bexar,Blanco,Borden,Bosque,Bowie,Brazoria,Brazos,Brewster,Briscoe,Brooks,Brown,Burleson,Burnet,Caldwell,Calhoun,Callahan,Cameron,Camp,Carson,Cass,Castro,Chambers,Cherokee,Childress,Clay,Cochran,Coke,Coleman,Collin,Collingsworth,Colorado,Comal,Comanche,Concho,Cooke,Coryell,Cottle,Crane,Crockett,Crosby,Culberson,Dallam,Dallas,Dawson,Deaf Smith,Delta,Denton,DeWitt,Dickens,Dimmit,Donley,Duval,Eastland,Ector,Edwards,El Paso,Ellis,Erath,Falls,Fannin,Fayette,Fisher,Floyd,Foard,Fort Bend,Franklin,Freestone,Frio,Gaines,Galveston,Garza,Gillespie,Glasscock,Goliad,Gonzales,Gray,Grayson,Gregg,Grimes,Guadalupe,Hale,Hall,Hamilton,Hansford,Hardeman,Hardin,Harris,Harrison,Hartley,Haskell,Hays,Hemphill,Henderson,Hidalgo,Hill,Hockley,Hood,Hopkins,Houston,Howard,Hudspeth,Hunt,Hutchinson,Irion,Jack,Jackson,Jasper,Jeff Davis,Jefferson,Jim Hogg,Jim Wells,Johnson,Jones,Karnes,Kaufman,Kendall,Kenedy,Kent,Kerr,Kimble,King,Kinney,Kleberg,Knox,La Salle,Lamar,Lamb,Lampasas,Lavaca,Lee,Leon,Liberty,Limestone,Lipscomb,Live Oak,Llano,Loving,Lubbock,Lynn,Madison,Marion,Martin,Mason,Matagorda,Maverick,McCulloch,McLennan,McMullen,Medina,Menard,Midland,Milam,Mills,Mitchell,Montague,Montgomery,Moore,Morris,Motley,Nacogdoches,Navarro,Newton,Nolan,Nueces,Ochiltree,Oldham,Orange,Palo Pinto,Panola,Parker,Parmer,Pecos,Polk,Potter,Presidio,Rains,Randall,Reagan,Real,Red River,Reeves,Refugio,Roberts,Robertson,Rockwall,Runnels,Rusk,Sabine,San Augustine,San Jacinto,San Patricio,San Saba,Schleicher,Scurry,Shackelford,Shelby,Sherman,Smith,Somervell,Starr,Stephens,Sterling,Stonewall,Sutton,Swisher,Tarrant,Taylor,Terrell,Terry,Throckmorton,Titus,Tom Green,Travis,Trinity,Tyler,Upshur,Upton,Uvalde,Val Verde,Van Zandt,Victoria,Walker,Waller,Ward,Washington,Webb,Wharton,Wheeler,Wichita,Wilbarger,Willacy,Williamson,Wilson,Winkler,Wise,Wood,Yoakum,Young,Zapata,Zavala',
    'Utah': 'Beaver,Box Elder,Cache,Carbon,Daggett,Davis,Duchesne,Emery,Garfield,Grand,Iron,Juab,Kane,Millard,Morgan,Piute,Rich,Salt Lake,San Juan,Sanpete,Sevier,Summit,Tooele,Uintah,Utah,Wasatch,Washington,Wayne,Weber',
    'Vermont': 'Addison,Bennington,Caledonia,Chittenden,Essex,Franklin,Grand Isle,Lamoille,Orange,Orleans,Rutland,Washington,Windham,Windsor',
    'Virginia': 'Accomack,Albemarle,Alleghany,Amelia,Amherst,Appomattox,Arlington,Augusta,Bath,Bedford,Bland,Botetourt,Brunswick,Buchanan,Buckingham,Campbell,Caroline,Carroll,Charles City,Charlotte,Chesterfield,Clarke,Craig,Culpeper,Cumberland,Dickenson,Dinwiddie,Essex,Fairfax,Fauquier,Floyd,Fluvanna,Franklin,Frederick,Giles,Gloucester,Goochland,Grayson,Greene,Greensville,Halifax,Hanover,Henrico,Henry,Highland,Isle of Wight,James City,King and Queen,King George,King William,Lancaster,Lee,Loudoun,Louisa,Lunenburg,Madison,Mathews,Mecklenburg,Middlesex,Montgomery,Nelson,New Kent,Northampton,Northumberland,Nottoway,Orange,Page,Patrick,Pittsylvania,Powhatan,Prince Edward,Prince George,Prince William,Pulaski,Rappahannock,Richmond,Roanoke,Rockbridge,Rockingham,Russell,Scott,Shenandoah,Smyth,Southampton,Spotsylvania,Stafford,Surry,Sussex,Tazewell,Warren,Washington,Westmoreland,Wise,Wythe,York',
    'Washington': 'Adams,Asotin,Benton,Chelan,Clallam,Clark,Columbia,Cowlitz,Douglas,Ferry,Franklin,Garfield,Grant,Grays Harbor,Island,Jefferson,King,Kitsap,Kittitas,Klickitat,Lewis,Lincoln,Mason,Okanogan,Pacific,Pend Oreille,Pierce,San Juan,Skagit,Skamania,Snohomish,Spokane,Stevens,Thurston,Wahkiakum,Walla Walla,Whatcom,Whitman,Yakima',
    'West Virginia': 'Barbour,Berkeley,Boone,Braxton,Brooke,Cabell,Calhoun,Clay,Doddridge,Fayette,Gilmer,Grant,Greenbrier,Hampshire,Hancock,Hardy,Harrison,Jackson,Jefferson,Kanawha,Lewis,Lincoln,Logan,Marion,Marshall,Mason,McDowell,Mercer,Mineral,Mingo,Monongalia,Monroe,Morgan,Nicholas,Ohio,Pendleton,Pleasants,Pocahontas,Preston,Putnam,Raleigh,Randolph,Ritchie,Roane,Summers,Taylor,Tucker,Tyler,Upshur,Wayne,Webster,Wetzel,Wirt,Wood,Wyoming',
    'Wisconsin': 'Adams,Ashland,Barron,Bayfield,Brown,Buffalo,Burnett,Calumet,Chippewa,Clark,Columbia,Crawford,Dane,Dodge,Door,Douglas,Dunn,Eau Claire,Florence,Fond du Lac,Forest,Grant,Green,Green Lake,Iowa,Iron,Jackson,Jefferson,Juneau,Kenosha,Kewaunee,La Crosse,Lafayette,Langlade,Lincoln,Manitowoc,Marathon,Marinette,Marquette,Menominee,Milwaukee,Monroe,Oconto,Oneida,Outagamie,Ozaukee,Pepin,Pierce,Polk,Portage,Price,Racine,Richland,Rock,Rusk,Sauk,Sawyer,Shawano,Sheboygan,St. Croix,Taylor,Trempealeau,Vernon,Vilas,Walworth,Washburn,Washington,Waukesha,Waupaca,Waushara,Winnebago,Wood',
    'Wyoming': 'Albany,Big Horn,Campbell,Carbon,Converse,Crook,Fremont,Goshen,Hot Springs,Johnson,Laramie,Lincoln,Natrona,Niobrara,Park,Platte,Sheridan,Sublette,Sweetwater,Teton,Uinta,Washakie,Weston',
    'District of Columbia': 'District of Columbia',
}

wb = openpyxl.load_workbook("SCG_CMA_System_v7.xlsx")
ws = wb["Counties Reference"]
ws.sheet_state = "hidden"

# Find next empty row
next_row = 3
while ws[f"A{next_row}"].value:
    next_row += 1

for state, county_list in counties.items():
    ws[f"A{next_row}"].value  = state
    ws[f"B{next_row}"].value  = county_list
    ws[f"A{next_row}"].font   = Font(size=9, name="Arial")
    ws[f"B{next_row}"].font   = Font(size=9, name="Arial")
    ws[f"A{next_row}"].border = thin()
    ws[f"B{next_row}"].border = thin()
    next_row += 1

wb.save("SCG_CMA_System_v7.xlsx")

last_data_row = next_row - 1
data_rows = last_data_row - 2   # subtract title row 1 + header row 2

print(f"Step 15e complete — {len(counties)} entries appended (rows 43-{last_data_row}).")
print(f"States/DC: {', '.join(counties.keys())}")
print(f"")
print(f"FINAL COUNT VERIFICATION:")
print(f"  Row 1:       Title banner")
print(f"  Row 2:       Headers (State | Counties)")
print(f"  Rows 3-{last_data_row}: Data ({data_rows} states/DC)")
print(f"  Total rows:  {last_data_row}")
print(f"  Data rows:   {data_rows} {'✓ CORRECT' if data_rows == 51 else '✗ EXPECTED 51'}")
