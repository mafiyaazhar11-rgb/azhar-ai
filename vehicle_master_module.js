// ============================================================
// VEHICLE MASTER MODULE — fleet registry used to fill in cost estimates
// when the dispatch file's own truck-type column is missing, by looking
// up the vehicle plate number instead.
// Mount with: require('./vehicle_master_module')(app, pool, requireAuth, requireRole, upload, auditLog, vehicleMasterMap);
// `vehicleMasterMap` is a plain object passed in by reference from server.js —
// this module populates it (never reassigns it) so server.js's synchronous
// dispatch-parsing code can read it without any DB calls mid-computation.
// ============================================================

const XLSX = require('xlsx');
const ExcelJS = require('exceljs');

// Embedded starting fleet list (from the vehicle master file provided) — this seeds the
// database automatically on first run, so no manual upload is needed to get going.
// The upload button still exists for later, whenever the fleet actually changes.
var VEHICLE_MASTER_SEED_DATA = [
  {vehicle_no:'DXB N 66510',department:'Consumer',model:'Canter 3 T',veh_type:'3T',vehicle_type_raw:'Chiller',weight:'3 T',adhoc:'DM CARD',partition_flag:'No'},
  {vehicle_no:'DXB N 67235',department:'Consumer',model:'Hino 10 T',veh_type:'10T',vehicle_type_raw:'Chiller',weight:'10 T',adhoc:'DM CARD',partition_flag:'No'},
  {vehicle_no:'DXB F 74669',department:'Consumer',model:'Canter 3 T',veh_type:'3T',vehicle_type_raw:'Chiller',weight:'3 T',adhoc:'DM CARD',partition_flag:'No'},
  {vehicle_no:'DXB O 96247',department:'Consumer',model:'Canter 3 T',veh_type:'3T',vehicle_type_raw:'Chiller',weight:'3 T',adhoc:'DM CARD',partition_flag:'No'},
  {vehicle_no:'DXB M 57467',department:'Consumer',model:'Canter 3 T',veh_type:'3T',vehicle_type_raw:'Chiller',weight:'3 T',adhoc:'DM CARD',partition_flag:'No'},
  {vehicle_no:'DXB F 74732',department:'Consumer',model:'Canter 3 T',veh_type:'3T',vehicle_type_raw:'Chiller',weight:'3 T',adhoc:'DM CARD',partition_flag:'No'},
  {vehicle_no:'AUH 18 18297',department:'Consumer',model:'Canter 4.2 T',veh_type:'4.2 T',vehicle_type_raw:'Chiller',weight:'4.2 T',adhoc:'DM CARD / FOOD WATCH',partition_flag:'No'},
  {vehicle_no:'DXB X 52466',department:'Consumer',model:'Canter 4.2 T',veh_type:'4.2 T',vehicle_type_raw:'Chiller',weight:'4.2 T',adhoc:'DM CARD',partition_flag:'No'},
  {vehicle_no:'DXB X 52467',department:'Consumer',model:'Canter 4.2 T',veh_type:'4.2 T',vehicle_type_raw:'Chiller',weight:'4.2 T',adhoc:'DM CARD',partition_flag:'No'},
  {vehicle_no:'DXB X 52468',department:'Consumer',model:'Fuso 12 ton',veh_type:'12T',vehicle_type_raw:'Chiller',weight:'12 T',adhoc:'DM CARD',partition_flag:'No'},
  {vehicle_no:'DXB X 52469',department:'Consumer',model:'Canter 4.2 T',veh_type:'4.2 T',vehicle_type_raw:'Chiller',weight:'4.2 T',adhoc:'DM CARD',partition_flag:'No'},
  {vehicle_no:'DXB X 52471',department:'Consumer',model:'Canter 4.2 T',veh_type:'4.2 T',vehicle_type_raw:'Chiller',weight:'4.2 T',adhoc:'DM CARD',partition_flag:'No'},
  {vehicle_no:'AUH 12 26577',department:'Consumer',model:'Canter 4.2 T',veh_type:'4.2 T',vehicle_type_raw:'Chiller',weight:'4.2 T',adhoc:'DM CARD / FOOD WATCH',partition_flag:'No'},
  {vehicle_no:'DXB I 92771',department:'Consumer',model:'Canter 3 T',veh_type:'3T',vehicle_type_raw:'Chiller',weight:'3 T',adhoc:'DM CARD',partition_flag:'No'},
  {vehicle_no:'DXB X 62975',department:'Consumer',model:'Canter 4.2 T',veh_type:'4.2 T',vehicle_type_raw:'Chiller',weight:'4.2 T',adhoc:'DM CARD',partition_flag:'No'},
  {vehicle_no:'AUH 16 74243',department:'Consumer',model:'Canter 4.2 T',veh_type:'4.2 T',vehicle_type_raw:'Chiller',weight:'4.2 T',adhoc:'DM CARD / FOOD WATCH',partition_flag:'No'},
  {vehicle_no:'DXB X 62978',department:'Consumer',model:'Canter 4.2 T',veh_type:'4.2 T',vehicle_type_raw:'Chiller',weight:'4.2 T',adhoc:'DM CARD',partition_flag:'No'},
  {vehicle_no:'DXB X 62979',department:'Consumer',model:'Canter 4.2 T',veh_type:'4.2 T',vehicle_type_raw:'Chiller',weight:'4.2 T',adhoc:'DM CARD',partition_flag:'No'},
  {vehicle_no:'DXB X 62982',department:'Consumer',model:'Canter 4.2 T',veh_type:'4.2 T',vehicle_type_raw:'Chiller',weight:'4.2 T',adhoc:'DM CARD',partition_flag:'No'},
  {vehicle_no:'DXB M 56601',department:'Consumer',model:'Canter 3 T',veh_type:'3T',vehicle_type_raw:'Chiller',weight:'3 T',adhoc:'DM CARD',partition_flag:'Yes'},
  {vehicle_no:'AUH 11 51038',department:'Consumer',model:'Canter 4.2 T',veh_type:'4.2 T',vehicle_type_raw:'Chiller',weight:'4.2 T',adhoc:'DM CARD / FOOD WATCH',partition_flag:'Yes'},
  {vehicle_no:'DXb J 38808',department:'Consumer',model:'Canter 4.2 T',veh_type:'4.2 T',vehicle_type_raw:'Chiller',weight:'4.2 T',adhoc:'DM CARD',partition_flag:'No'},
  {vehicle_no:'DXB D 50149',department:'Consumer',model:'Canter 4.2 T',veh_type:'4.2 T',vehicle_type_raw:'Chiller',weight:'4.2 T',adhoc:'DM CARD',partition_flag:'Yes'},
  {vehicle_no:'AUH 5 12836',department:'Consumer',model:'Canter 4.2 T',veh_type:'4.2 T',vehicle_type_raw:'Chiller',weight:'4.2 T',adhoc:'DM CARD / FOOD WATCH',partition_flag:'Yes'},
  {vehicle_no:'DXB W 50238',department:'Consumer',model:'Canter 4.2 T',veh_type:'4.2 T',vehicle_type_raw:'Chiller',weight:'4.2 T',adhoc:'DM CARD',partition_flag:'Yes'},
  {vehicle_no:'DXB W 50242',department:'Consumer',model:'Canter 4.2 T',veh_type:'4.2 T',vehicle_type_raw:'Chiller',weight:'4.2 T',adhoc:'DM CARD',partition_flag:'Yes'},
  {vehicle_no:'AUH 20 99641',department:'Consumer',model:'Canter 4.2 T',veh_type:'4.2 T',vehicle_type_raw:'Chiller',weight:'4.2 T',adhoc:'DM CARD / FOOD WATCH',partition_flag:'Yes'},
  {vehicle_no:'DXB D 70426',department:'Consumer',model:'Canter 4.2 T',veh_type:'4.2 T',vehicle_type_raw:'Chiller',weight:'4.2 T',adhoc:'DM CARD',partition_flag:'Yes'},
  {vehicle_no:'DXB H 93746',department:'Consumer',model:'Canter 4.2 T',veh_type:'4.2 T',vehicle_type_raw:'Chiller',weight:'4.2 T',adhoc:'DM CARD',partition_flag:'Yes'},
  {vehicle_no:'DXB H 93956',department:'Consumer',model:'Canter 4.2 T',veh_type:'4.2 T',vehicle_type_raw:'Chiller',weight:'4.2 T',adhoc:'DM CARD',partition_flag:'Yes'},
  {vehicle_no:'DXB W 68840',department:'Consumer',model:'Canter 4.2 T',veh_type:'4.2 T',vehicle_type_raw:'Chiller',weight:'4.2 T',adhoc:'DM CARD',partition_flag:'Yes'},
  {vehicle_no:'DXB L 87119',department:'Consumer',model:'Canter 4.2 T',veh_type:'4.2 T',vehicle_type_raw:'Chiller',weight:'4.2 T',adhoc:'DM CARD',partition_flag:'Yes'},
  {vehicle_no:'DXB R 46983',department:'Consumer',model:'Urvan 2.4 L',veh_type:'1T',vehicle_type_raw:'Chiller',weight:'1 T',adhoc:'DM CARD',partition_flag:'Yes'},
  {vehicle_no:'DXB H 48432',department:'Consumer',model:'Canter 4.2 T',veh_type:'4.2 T',vehicle_type_raw:'Chiller',weight:'4.2 T',adhoc:'DM CARD',partition_flag:'Yes'},
  {vehicle_no:'AUH 17 91349',department:'Consumer',model:'Hiace Van',veh_type:'1T',vehicle_type_raw:'Chiller',weight:'1 T',adhoc:'DM CARD / FOOD WATCH',partition_flag:'No'},
  {vehicle_no:'DXB Y 53725',department:'Consumer',model:'Hiace Van',veh_type:'1T',vehicle_type_raw:'Chiller',weight:'1 T',adhoc:'DM CARD',partition_flag:'No'},
  {vehicle_no:'DXB E 20465',department:'Consumer',model:'Canter 4.2 T',veh_type:'4.2 T',vehicle_type_raw:'Chiller',weight:'4.2 T',adhoc:'DM CARD',partition_flag:'Yes'},
  {vehicle_no:'DXB P 47964',department:'Consumer',model:'Canter 4.2 T',veh_type:'4.2 T',vehicle_type_raw:'Chiller',weight:'4.2 T',adhoc:'DM CARD',partition_flag:'Yes'},
  {vehicle_no:'DXB E 32576',department:'Consumer',model:'Canter 4.2 T',veh_type:'4.2 T',vehicle_type_raw:'Chiller',weight:'4.2 T',adhoc:'DM CARD',partition_flag:'No'},
  {vehicle_no:'DXB E 38248',department:'Consumer',model:'Canter 4.2 T',veh_type:'4.2 T',vehicle_type_raw:'Chiller',weight:'4.2 T',adhoc:'DM CARD',partition_flag:'Yes'},
  {vehicle_no:'DXB E 69044',department:'Consumer',model:'Canter 4.2 T',veh_type:'4.2 T',vehicle_type_raw:'Chiller',weight:'4.2 T',adhoc:'DM CARD',partition_flag:'No'},
  {vehicle_no:'DXB E 13041',department:'Consumer',model:'Fuso 12 ton',veh_type:'12T',vehicle_type_raw:'Chiller',weight:'12 T',adhoc:'DM CARD',partition_flag:'No'},
  {vehicle_no:'AUH 21 60587',department:'Consumer',model:'Fuso 12 ton',veh_type:'12T',vehicle_type_raw:'Chiller',weight:'12 T',adhoc:'DM CARD / FOOD WATCH',partition_flag:'No'},
  {vehicle_no:'DXB E 48971',department:'Consumer',model:'Canter 4.2 T',veh_type:'4.2 T',vehicle_type_raw:'Chiller',weight:'4.2 T',adhoc:'DM CARD',partition_flag:'No'},
  {vehicle_no:'DXB J 83979',department:'Consumer',model:'Urvan 2.4 L',veh_type:'1T',vehicle_type_raw:'Chiller',weight:'1 T',adhoc:'DM CARD',partition_flag:'No'},
  {vehicle_no:'DXB B 54498',department:'Consumer',model:'Isuzu NPR 4.2 ton',veh_type:'4.2 T',vehicle_type_raw:'Ambient',weight:'4.2 T',adhoc:'DM CARD',partition_flag:'Yes'},
  {vehicle_no:'DXB D 98207',department:'Consumer',model:'Nissan Urvan',veh_type:'1T',vehicle_type_raw:'Frozen',weight:null,adhoc:'DM CARD',partition_flag:'Yes'},
  {vehicle_no:'AUH 13 60787',department:'Consumer',model:'Nissan Urvan',veh_type:'1T',vehicle_type_raw:'Frozen',weight:null,adhoc:'DM CARD / FOOD WATCH',partition_flag:'Yes'},
  {vehicle_no:'AUH 18 25896',department:'Consumer',model:'Isuzu NPR 4.2 ton',veh_type:'4.2 T',vehicle_type_raw:'Chiller',weight:'4.2',adhoc:'DM CARD',partition_flag:'Yes'},
  {vehicle_no:'AUH 18 25848',department:'Consumer',model:'Isuzu NPR 4.2 ton',veh_type:'4.2 T',vehicle_type_raw:'Chiller',weight:'4.2',adhoc:'DM CARD',partition_flag:'Yes'},
  {vehicle_no:'AUH 18 24935',department:'Consumer',model:'Isuzu NPR 4.2 ton',veh_type:'4.2 T',vehicle_type_raw:'Chiller',weight:'4.2',adhoc:'DM CARD',partition_flag:'Yes'},
  {vehicle_no:'AUH 18 24819',department:'Consumer',model:'Isuzu NPR 4.2 ton',veh_type:'4.2 T',vehicle_type_raw:'Chiller',weight:'4.2',adhoc:'DM CARD',partition_flag:'Yes'},
  {vehicle_no:'AUH 18 24795',department:'Consumer',model:'Isuzu NPR 4.2 ton',veh_type:'4.2 T',vehicle_type_raw:'Chiller',weight:'4.2',adhoc:'DM CARD',partition_flag:'Yes'},
  {vehicle_no:'AUH 5 63359',department:'Consumer',model:'Isuzu NPR 4.2 ton',veh_type:'4.2 T',vehicle_type_raw:'Chiller',weight:'4.2',adhoc:'DM CARD / FOOD WATCH',partition_flag:'YES'},
  {vehicle_no:'AUH 5 63107',department:'Consumer',model:'Isuzu NPR 4.2 ton',veh_type:'4.2 T',vehicle_type_raw:'Chiller',weight:'4.2',adhoc:'DM CARD / FOOD WATCH',partition_flag:'YES'},
  {vehicle_no:'AUH 5 62842',department:'Consumer',model:'Isuzu NPR 4.2 ton',veh_type:'4.2 T',vehicle_type_raw:'Chiller',weight:'4.2',adhoc:'DM CARD / FOOD WATCH',partition_flag:'YES'},
  {vehicle_no:'AUH 5 62921',department:'Consumer',model:'Isuzu NPR 4.2 ton',veh_type:'4.2 T',vehicle_type_raw:'Chiller',weight:'4.2',adhoc:'DM CARD / FOOD WATCH',partition_flag:'YES'},
  {vehicle_no:'AUH 5 63381',department:'Consumer',model:'Isuzu NPR 4.2 ton',veh_type:'4.2 T',vehicle_type_raw:'Chiller',weight:'4.2',adhoc:'DM CARD / FOOD WATCH',partition_flag:'YES'},
  {vehicle_no:'DXB W 50236',department:'Consumer / Frozen',model:'Canter 4.2 T',veh_type:'4.2 T',vehicle_type_raw:'Chiller',weight:'4.2 T',adhoc:'DM CARD',partition_flag:'Yes'},
  {vehicle_no:'DXB B 54486',department:'Consumer / Frozen',model:'Isuzu NPR 4.2 ton',veh_type:'4.2 T',vehicle_type_raw:'Frozen',weight:'4.2 T',adhoc:'DM CARD',partition_flag:'Yes'},
  {vehicle_no:'AUH 7 27860',department:'Consumer / Frozen',model:'Isuzu NPR 4.2 ton',veh_type:'4.2 T',vehicle_type_raw:'Frozen',weight:'4.2 T',adhoc:'DM CARD / FOOD WATCH',partition_flag:'Yes'},
  {vehicle_no:'DXB B 54437',department:'Consumer / Frozen',model:'Isuzu NPR 4.2 ton',veh_type:'4.2 T',vehicle_type_raw:'Frozen',weight:'4.2 T',adhoc:'DM CARD',partition_flag:'Yes'},
  {vehicle_no:'DXB B 54438',department:'Consumer / Frozen',model:'Isuzu NPR 4.2 ton',veh_type:'4.2 T',vehicle_type_raw:'Frozen',weight:'4.2 T',adhoc:'DM CARD',partition_flag:'Yes'},
  {vehicle_no:'DXB B 54471',department:'Consumer / Frozen',model:'Isuzu NPR 4.2 ton',veh_type:'4.2 T',vehicle_type_raw:'Frozen',weight:'4.2 T',adhoc:'DM CARD',partition_flag:'Yes'},
  {vehicle_no:'DXB S 42163',department:'Consumer / Frozen',model:'Isuzu NPR 4.2 ton',veh_type:'4.2 T',vehicle_type_raw:'Frozen',weight:'4.2',adhoc:'DM CARD',partition_flag:'Yes'},
  {vehicle_no:'AUH 18 73985',department:'Consumer / Frozen',model:'Isuzu NPR 4.2 ton',veh_type:'4.2 T',vehicle_type_raw:'Frozen',weight:'4.2',adhoc:'DM CARD / FOOD WATCH',partition_flag:'Yes'},
  {vehicle_no:'AUH 18 70998',department:'Consumer / Frozen',model:'Isuzu NPR 4.2 ton',veh_type:'4.2 T',vehicle_type_raw:'Frozen',weight:'4.2',adhoc:'DM CARD / FOOD WATCH',partition_flag:'Yes'},
  {vehicle_no:'AUH 18 76372',department:'Consumer / Frozen',model:'Isuzu NPR 4.2 ton',veh_type:'4.2 T',vehicle_type_raw:'Frozen',weight:'4.2',adhoc:'DM CARD / FOOD WATCH',partition_flag:'Yes'},
  {vehicle_no:'AUH 18 73914',department:'Consumer / Frozen',model:'Isuzu NPR 4.2 ton',veh_type:'4.2 T',vehicle_type_raw:'Frozen',weight:'4.2',adhoc:'DM CARD / FOOD WATCH',partition_flag:'Yes'},
  {vehicle_no:'AUH 21 64731',department:'Consumer / Frozen',model:'Isuzu NPR 3 ton',veh_type:'3T',vehicle_type_raw:'Freezer',weight:'3T',adhoc:'DM CARD / FOOD WATCH',partition_flag:'YES'},
  {vehicle_no:'AUH 21 64732',department:'Consumer / Frozen',model:'Isuzu NPR 3 ton',veh_type:'3T',vehicle_type_raw:'Freezer',weight:'3T',adhoc:'DM CARD / FOOD WATCH',partition_flag:'YES'},
  {vehicle_no:'DXB K 86316',department:'Consumer / Frozen/zooz',model:'Mitsubishi Canter',veh_type:'3T',vehicle_type_raw:'Frozen',weight:'-',adhoc:'DM CARD',partition_flag:'No'},
  {vehicle_no:'DXB K 86319',department:'Consumer / Frozen/zooz',model:'Mitsubishi Canter',veh_type:'3T',vehicle_type_raw:'Frozen',weight:'-',adhoc:'DM CARD',partition_flag:'No'},
  {vehicle_no:'DXB F 75568',department:'Consumer / Stand By',model:'Canter 3 T',veh_type:'3T',vehicle_type_raw:'Chiller',weight:'3 T',adhoc:'DM CARD',partition_flag:'No'},
  {vehicle_no:'AUH 14 97923',department:'Consumer / ZOOZ',model:'MITSUBISHI CANTER',veh_type:'3T',vehicle_type_raw:'Chiller',weight:'3T',adhoc:'DM CARD / FOOD WATCH',partition_flag:'NO'},
  {vehicle_no:'AUH 13 71022',department:'Consumer / ZOOZ',model:'MITSUBISHI CANTER',veh_type:'3T',vehicle_type_raw:'Chiller',weight:'3T',adhoc:'DM CARD / FOOD WATCH',partition_flag:'NO'},
  {vehicle_no:'DXB L 90872',department:'Consumer / ZOOZ',model:'MITSUBISHI CANTER',veh_type:'3T',vehicle_type_raw:'Chiller',weight:'3T',adhoc:'DM CARD / FOOD WATCH',partition_flag:'NO'},
  {vehicle_no:'AUH 14 17609',department:'Consumer / ZOOZ',model:'MITSUBISHI CANTER',veh_type:'3T',vehicle_type_raw:'Chiller',weight:'3T',adhoc:'DM CARD / FOOD WATCH',partition_flag:'NO'},
  {vehicle_no:'DXB H 51931',department:'Consumer / ZOOZ',model:'MITSUBISHI CANTER',veh_type:'3T',vehicle_type_raw:'chiller',weight:'3T',adhoc:'DM CARD',partition_flag:'n/a'},
  {vehicle_no:'DXB G 21573',department:'Consumer / ZOOZ',model:'MITSUBISHI CANTER',veh_type:'3T',vehicle_type_raw:'chiller',weight:'3T',adhoc:'Dry Food',partition_flag:'n/a'},
  {vehicle_no:'DXB G 75508',department:'Consumer / ZOOZ',model:'MITSUBISHI CANTER',veh_type:'3T',vehicle_type_raw:'chiller',weight:'3T',adhoc:'Dry Food',partition_flag:'n/a'},
  {vehicle_no:'DXB W 82364',department:'Consumer / ZOOZ',model:'TOYOTA HIACE',veh_type:'1T',vehicle_type_raw:'Chiller',weight:'1T',adhoc:'DM CARD chiller',partition_flag:'n/a'},
  {vehicle_no:'DXB Y 69707',department:'Consumer / ZOOZ',model:'TOYOTA HIACE',veh_type:'1T',vehicle_type_raw:'Chiller',weight:'1T',adhoc:'DM CARD chiller',partition_flag:'n/a'},
  {vehicle_no:'DXB H 29683',department:'Consumer / ZOOZ',model:'MITSUBISHI CANTER',veh_type:'3T',vehicle_type_raw:'Chiller',weight:'3T',adhoc:'DM CARD chiller',partition_flag:'n/a'},
  {vehicle_no:'AUH 23 91439',department:'Consumer/Electric',model:'SANY LIGHT',veh_type:'3T',vehicle_type_raw:'Freezer',weight:'3T',adhoc:'DM CARD',partition_flag:'NO'},
  {vehicle_no:'AUH 23 91440',department:'Consumer/Electric',model:'SANY LIGHT',veh_type:'3T',vehicle_type_raw:'Freezer',weight:'3T',adhoc:'DM CARD',partition_flag:'NO'},
  {vehicle_no:'DXB CC 88298',department:'DAHC',model:'Mitsubishi Canter',veh_type:'3T',vehicle_type_raw:'Chiller',weight:'4.2',adhoc:'DM CARD',partition_flag:'No'},
  {vehicle_no:'DXB CC 88301',department:'DAHC',model:'Mitsubishi Canter',veh_type:'3T',vehicle_type_raw:'Chiller',weight:'4.2',adhoc:'DM CARD',partition_flag:'No'},
  {vehicle_no:'DXB CC 88302',department:'DAHC',model:'Toyota Lite ACE',veh_type:'1T',vehicle_type_raw:'Chiller',weight:'-',adhoc:'DM CARD',partition_flag:'No'},
  {vehicle_no:'DXB CC 88304',department:'DAHC',model:'Toyota Lite ACE',veh_type:'1T',vehicle_type_raw:'Chiller',weight:'-',adhoc:'DM CARD',partition_flag:'No'},
  {vehicle_no:'DXB CC 88306',department:'DAHC',model:'Toyota Lite ACE',veh_type:'1T',vehicle_type_raw:'Chiller',weight:'-',adhoc:'DM CARD',partition_flag:'No'},
  {vehicle_no:'DXB CC 81648',department:'DAHC',model:'Mitsubishi Canter',veh_type:'3T',vehicle_type_raw:'Chiller',weight:'-',adhoc:'DM CARD',partition_flag:'No'},
  {vehicle_no:'DXB CC 81649',department:'DAHC',model:'Mitsubishi Canter',veh_type:'3T',vehicle_type_raw:'Chiller',weight:'-',adhoc:'DM CARD',partition_flag:'No'},
  {vehicle_no:'DXB CC 81650',department:'DAHC',model:'Mitsubishi Canter',veh_type:'3T',vehicle_type_raw:'Chiller',weight:'-',adhoc:'DM CARD',partition_flag:'No'},
  {vehicle_no:'DXB CC 85608',department:'DAHC',model:'Isuzu NPR 3 ton',veh_type:'3T',vehicle_type_raw:'Chiller',weight:'3T',adhoc:'DM CARD',partition_flag:'Yes'},
  {vehicle_no:'DXB CC 75808',department:'DAHC',model:'Isuzu NPR 3 ton',veh_type:'3T',vehicle_type_raw:'Chiller',weight:'3T',adhoc:'DM CARD',partition_flag:'Yes'},
  {vehicle_no:'DXB N 77586',department:'DCF',model:'Toyota Hiace',veh_type:'1T',vehicle_type_raw:'Chiller',weight:'-',adhoc:'DM CARD',partition_flag:'No'},
  {vehicle_no:'DXB U 78248',department:'Fitness',model:'Canter 4.2 T',veh_type:'4.2 T',vehicle_type_raw:'Dry',weight:'4.2 T',adhoc:'-',partition_flag:'No'},
  {vehicle_no:'DXB U 78249',department:'Fitness',model:'Canter 4.2 T',veh_type:'4.2 T',vehicle_type_raw:'Dry',weight:'4.2 T',adhoc:'-',partition_flag:'No'},
  {vehicle_no:'DXB U 78250',department:'Fitness',model:'Canter 4.2 T',veh_type:'4.2 T',vehicle_type_raw:'Dry',weight:'4.2 T',adhoc:'-',partition_flag:'No'},
  {vehicle_no:'DXB BB 64355',department:'Fitness',model:'HINO 300',veh_type:'10T',vehicle_type_raw:'Open Pick-up',weight:'4.2 T',adhoc:'-',partition_flag:'No'},
  {vehicle_no:'DXB J 82930',department:'Fleet/Pharma',model:'HI',veh_type:'Car',vehicle_type_raw:'Car',weight:'1 T',adhoc:'-',partition_flag:'No'},
  {vehicle_no:'DXB S 40087',department:'General',model:'Avanza 1.5',veh_type:'1T',vehicle_type_raw:'Other',weight:'staff',adhoc:'-',partition_flag:'No'},
  {vehicle_no:'DXB S 22045',department:'Hospital Consignment',model:'ATTRAGE',veh_type:'Car',vehicle_type_raw:'Other',weight:'Car',adhoc:'-',partition_flag:'No'},
  {vehicle_no:'DXB Z 93849',department:'NMC',model:'Isuzu NPR 4.2 ton',veh_type:'4.2T',vehicle_type_raw:'Chiller',weight:'4.2',adhoc:'DM CARD',partition_flag:'YES'},
  {vehicle_no:'DXB Z 85818',department:'NMC',model:'Isuzu NPR 4.2 ton',veh_type:'4.2T',vehicle_type_raw:'Chiller',weight:'4.2',adhoc:'DM CARD',partition_flag:'YES'},
  {vehicle_no:'AUH 22 42156',department:'NMC',model:'Isuzu NPR 4.2 ton',veh_type:'4.2T',vehicle_type_raw:'Chiller',weight:'4.2',adhoc:'DM CARD / FOOD WATCH',partition_flag:'YES'},
  {vehicle_no:'AUH 22 42157',department:'NMC',model:'Isuzu NPR 4.2 ton',veh_type:'4.2T',vehicle_type_raw:'Chiller',weight:'4.2',adhoc:'DM CARD / FOOD WATCH',partition_flag:'YES'},
  {vehicle_no:'AUH 22 40858',department:'NMC',model:'Isuzu NPR 4.2 ton',veh_type:'4.2T',vehicle_type_raw:'Chiller',weight:'4.2',adhoc:'DM CARD / FOOD WATCH',partition_flag:'YES'},
  {vehicle_no:'AUH 16 63542',department:'Pharma',model:'Canter 4.2 T',veh_type:'4.2 T',vehicle_type_raw:'Chiller',weight:'4.2 T',adhoc:'DM CARD / FOOD WATCH',partition_flag:'No'},
  {vehicle_no:'DXB O 96244',department:'Pharma',model:'Canter 3 T',veh_type:'3T',vehicle_type_raw:'Chiller',weight:'3 T',adhoc:'DM CARD',partition_flag:'No'},
  {vehicle_no:'AUH 20 99642',department:'Pharma',model:'Canter 3 T',veh_type:'3T',vehicle_type_raw:'Chiller',weight:'3 T',adhoc:'DM CARD / FOOD WATCH',partition_flag:'No'},
  {vehicle_no:'DXB I 94801',department:'Pharma',model:'Canter 3 T',veh_type:'3T',vehicle_type_raw:'Chiller',weight:'3 T',adhoc:'DM CARD',partition_flag:'Yes'},
  {vehicle_no:'AUH 20 63924',department:'Pharma',model:'Canter 3 T',veh_type:'3T',vehicle_type_raw:'Chiller',weight:'3 T',adhoc:'Food Watch',partition_flag:'No'},
  {vehicle_no:'DXB R 23159',department:'Pharma',model:'Urvan 2.4 L',veh_type:'1T',vehicle_type_raw:'Chiller',weight:'1 T',adhoc:'DM CARD',partition_flag:'No'},
  {vehicle_no:'DXB J 75017',department:'Pharma',model:'Urvan 2.4 L',veh_type:'1T',vehicle_type_raw:'Chiller',weight:'1 T',adhoc:'DM CARD',partition_flag:'No'},
  {vehicle_no:'DXB J 81651',department:'Pharma',model:'Urvan 2.4 L',veh_type:'1T',vehicle_type_raw:'Chiller',weight:'1 T',adhoc:'DM CARD',partition_flag:'No'},
  {vehicle_no:'DXB W 50232',department:'Pharma',model:'Canter 4.2 T',veh_type:'4.2 T',vehicle_type_raw:'Chiller',weight:'4.2 T',adhoc:'DM CARD',partition_flag:'Yes'},
  {vehicle_no:'DXB H 47655',department:'Pharma',model:'Canter 4.2 T',veh_type:'4.2 T',vehicle_type_raw:'Chiller',weight:'4.2 T',adhoc:'DM CARD',partition_flag:'Yes'},
  {vehicle_no:'AUH 12 26793',department:'Pharma',model:'Canter 4.2 T',veh_type:'4.2 T',vehicle_type_raw:'Chiller',weight:'4.2 T',adhoc:'DM CARD / FOOD WATCH',partition_flag:'Yes'},
  {vehicle_no:'AUH 14 24013',department:'Pharma',model:'Canter 4.2 T',veh_type:'4.2 T',vehicle_type_raw:'Chiller',weight:'4.2 T',adhoc:'DM CARD / FOOD WATCH',partition_flag:'Yes'},
  {vehicle_no:'DXB J 87205',department:'Pharma',model:'Urvan 2.4 L',veh_type:'1T',vehicle_type_raw:'Chiller',weight:'1 T',adhoc:'DM CARD',partition_flag:'No'},
  {vehicle_no:'DXB P 51712',department:'Pharma',model:'Canter 4.2 T',veh_type:'4.2 T',vehicle_type_raw:'Chiller',weight:'4.2 T',adhoc:'DM CARD',partition_flag:'Yes'},
  {vehicle_no:'DXB D 82157',department:'Pharma',model:'Fuso 12 ton',veh_type:'12T',vehicle_type_raw:'Food  Chiller',weight:'12 T',adhoc:'DM CARD',partition_flag:'No'},
  {vehicle_no:'DXB J 59323',department:'pharma',model:'Isuzu NPR 4.2 ton',veh_type:'4.2 T',vehicle_type_raw:'CHILLER',weight:'4.2 T',adhoc:'DM CARD',partition_flag:'No'},
  {vehicle_no:'DXB J 52061',department:'pharma',model:'Isuzu NPR 4.2 ton',veh_type:'4.2 T',vehicle_type_raw:'CHILLER',weight:'4.2 T',adhoc:'DM CARD',partition_flag:'No'},
  {vehicle_no:'DXB J 60138',department:'pharma',model:'Isuzu NPR 4.2 ton',veh_type:'4.2 T',vehicle_type_raw:'CHILLER',weight:'4.2 T',adhoc:'DM CARD',partition_flag:'No'},
  {vehicle_no:'DXB J 27188',department:'pharma',model:'Isuzu NPR 4.2 ton',veh_type:'4.2 T',vehicle_type_raw:'CHILLER',weight:'4.2 T',adhoc:'DM CARD',partition_flag:'No'},
  {vehicle_no:'DXB J 58076',department:'pharma',model:'Isuzu NPR 4.2 ton',veh_type:'4.2 T',vehicle_type_raw:'CHILLER',weight:'4.2 T',adhoc:'DM CARD',partition_flag:'No'},
  {vehicle_no:'DXB B 54436',department:'pharma',model:'Isuzu NPR 4.2 ton',veh_type:'4.2 T',vehicle_type_raw:'Ambient',weight:'4.2 T',adhoc:'DM CARD',partition_flag:'Yes'},
  {vehicle_no:'AUH 5 65826',department:'Pharma',model:'Mitsubishi Fuso 12T',veh_type:'12T',vehicle_type_raw:'Ambient',weight:'12 T',adhoc:'DM CARD / FOOD WATCH',partition_flag:'Yes'},
  {vehicle_no:'DXB O 89042',department:'Pharma',model:'Isuzu NPR 4.2 ton',veh_type:'4.2 T',vehicle_type_raw:'Chiller',weight:'4.2',adhoc:'DM CARD',partition_flag:'Yes'},
  {vehicle_no:'DXB O 90419',department:'Pharma',model:'Isuzu NPR 4.2 ton',veh_type:'4.2 T',vehicle_type_raw:'Chiller',weight:'4.2',adhoc:'DM CARD',partition_flag:'Yes'},
  {vehicle_no:'AUH 18 73108',department:'Pharma',model:'Isuzu NPR 4.2 ton',veh_type:'4.2 T',vehicle_type_raw:'Chiller',weight:'4.2',adhoc:'DM CARD / FOOD WATCH',partition_flag:'Yes'},
  {vehicle_no:'AUH 13 35029',department:'Pharma',model:'Isuzu NPR 4.2 ton',veh_type:'4.2 T',vehicle_type_raw:'Chiller',weight:'4.2',adhoc:'DM CARD / FOOD WATCH',partition_flag:'Yes'},
  {vehicle_no:'AUH 13 27970',department:'Pharma',model:'Isuzu NPR 4.2 ton',veh_type:'4.2 T',vehicle_type_raw:'Chiller',weight:'4.2',adhoc:'DM CARD / FOOD WATCH',partition_flag:'Yes'},
  {vehicle_no:'AUH 12 22703',department:'Pharma / Dental',model:'Van',veh_type:'1T',vehicle_type_raw:'Chiller',weight:'1 T',adhoc:'DM CARD / FOOD WATCH',partition_flag:'No'},
  {vehicle_no:'AUH 22 89336',department:'Pharma / E-Com',model:'Toyota Hiace',veh_type:'1T',vehicle_type_raw:'Chiller',weight:'1 T',adhoc:'DM CARD / FOOD WATCH',partition_flag:'No'},
  {vehicle_no:'AUH 22 89335',department:'Pharma / E-Com',model:'Toyota Hiace',veh_type:'1T',vehicle_type_raw:'Chiller',weight:'1 T',adhoc:'DM CARD / FOOD WATCH',partition_flag:'No'},
  {vehicle_no:'AUH 11 36596',department:'Pharma / E-Com',model:'Toyota Hiace',veh_type:'1T',vehicle_type_raw:'Chiller',weight:'1 T',adhoc:'DM CARD / FOOD WATCH',partition_flag:'No'},
  {vehicle_no:'DXB EE 71791',department:'Pharma / E-Com',model:'Toyota Hiace',veh_type:'1T',vehicle_type_raw:'Chiller',weight:'1 T',adhoc:'DM CARD',partition_flag:'No'},
  {vehicle_no:'DXB EE 71790',department:'Pharma / E-Com',model:'Toyota Hiace',veh_type:'1T',vehicle_type_raw:'Chiller',weight:'1 T',adhoc:'DM CARD',partition_flag:'No'},
  {vehicle_no:'DXB EE 77549',department:'Pharma / E-Com',model:'Toyota Hiace',veh_type:'1T',vehicle_type_raw:'Chiller',weight:'1 T',adhoc:'DM CARD',partition_flag:'No'},
  {vehicle_no:'AUH 12 21393',department:'Pharma / Open Pickup',model:'Canter 3 T',veh_type:'3T',vehicle_type_raw:'Open Pick-up',weight:'3 T',adhoc:'-',partition_flag:'No'},
  {vehicle_no:'DXB CC 88305',department:'Pharma / Salon',model:'Toyota Lite ACE',veh_type:'1T',vehicle_type_raw:'Chiller',weight:'-',adhoc:'DM CARD',partition_flag:'No'},
  {vehicle_no:'DXB CC 88307',department:'Pharma / Salon',model:'Toyota Lite ACE',veh_type:'1T',vehicle_type_raw:'Chiller',weight:'-',adhoc:'DM CARD',partition_flag:'No'},
  {vehicle_no:'DXB G 30334',department:'Pharma / ZOOZ',model:'NISSAN URVAN',veh_type:'1T',vehicle_type_raw:'Chiller',weight:'1T',adhoc:'DM CARD chiller',partition_flag:'n/a'},
  {vehicle_no:'AUH 11 39705',department:'Pharma/ Lab Loaner',model:'Van',veh_type:'Car',vehicle_type_raw:'Other',weight:'Car',adhoc:'-',partition_flag:'No'},
  {vehicle_no:'AUH 20 86849',department:'Pharma/ Lab Loaner',model:'Van',veh_type:'Car',vehicle_type_raw:'van',weight:'1 T',adhoc:'-',partition_flag:'No'},
  {vehicle_no:'DXB V 98003',department:'Staff transport / Consumer',model:'Bus',veh_type:'Bus',vehicle_type_raw:'Other',weight:'staff',adhoc:'-',partition_flag:'No'},
  {vehicle_no:'DXB O 28615',department:'Staff transport / Consumer',model:'ASHOK LEYLAND',veh_type:'BUS',vehicle_type_raw:null,weight:null,adhoc:'-',partition_flag:'NO'},
  {vehicle_no:'DXB S 75141',department:'Staff transport / Pharma',model:'14s',veh_type:'Car',vehicle_type_raw:'Other',weight:'staff',adhoc:'Speed Limiter',partition_flag:'No'},
  {vehicle_no:'DXB H 25096',department:'Staff transport / Pharma',model:'bus',veh_type:'Bus',vehicle_type_raw:'Other',weight:'staff',adhoc:'-',partition_flag:'No'},
  {vehicle_no:'DXB Q 49762',department:'Staff transport / Pharma',model:'Hiace Seater',veh_type:'Car',vehicle_type_raw:'Other',weight:'staff',adhoc:'-',partition_flag:'No'},
  {vehicle_no:'DXB W 37762',department:'Staff transport / Pharma',model:'BUS',veh_type:'Bus',vehicle_type_raw:'Other',weight:'staff',adhoc:'-',partition_flag:'No'},
  {vehicle_no:'DXB O 31175',department:'Staff transport / Pharma',model:'ASHOK LEYLAND',veh_type:'BUS',vehicle_type_raw:null,weight:null,adhoc:'-',partition_flag:'NO'},
  {vehicle_no:'DXB Q 52482',department:'Zooz',model:'Toyota Hiace',veh_type:'1T',vehicle_type_raw:'Chiller',weight:'-',adhoc:'DM CARD',partition_flag:'No'},
  {vehicle_no:'DXB Y 58545',department:'Zooz',model:'Toyota Hiace',veh_type:'1T',vehicle_type_raw:'Chiller',weight:'-',adhoc:'DM CARD',partition_flag:'n/a'},
  {vehicle_no:'DXB S 85415',department:'Zooz',model:'Toyota Hiace',veh_type:'1T',vehicle_type_raw:'Chiller',weight:'-',adhoc:'DM CARD',partition_flag:'No'},
  {vehicle_no:'DXB G 62413',department:'Zooz',model:'Mitsubishi Canter',veh_type:'3T',vehicle_type_raw:'Chiller',weight:'-',adhoc:'DM CARD',partition_flag:'No'},
  {vehicle_no:'DXB G 22415',department:'Zooz',model:'Mitsubishi Canter',veh_type:'3T',vehicle_type_raw:'Chiller',weight:'-',adhoc:'DM CARD',partition_flag:'No'},
  {vehicle_no:'DXB Q 52483',department:'Zooz',model:'Toyota Hiace',veh_type:'1T',vehicle_type_raw:'Chiller',weight:'-',adhoc:'DM CARD',partition_flag:'No'},
  {vehicle_no:'DXB G 66947',department:'Zooz',model:'Mitsubishi Canter',veh_type:'3T',vehicle_type_raw:'Chiller',weight:'3T',adhoc:'DM CARD',partition_flag:'NO'},
  {vehicle_no:'DXB S 84934',department:'Zooz',model:'Toyota Hiace',veh_type:'1T',vehicle_type_raw:'Chiller',weight:'1T',adhoc:'DM CARD',partition_flag:'YES'},
  {vehicle_no:'DXB G 21709',department:'Zooz',model:'Mitsubishi Canter',veh_type:'3T',vehicle_type_raw:'Chiller',weight:'3T',adhoc:'DM CARD',partition_flag:'NO'},
  {vehicle_no:'DXB S 85411',department:'Zooz',model:'Toyota Hiace',veh_type:'1T',vehicle_type_raw:'Chiller',weight:'1T',adhoc:'DM CARD',partition_flag:'YES'},
  {vehicle_no:'AUH 7 27889',department:'Zooz',model:'MITSUBISHI CANTER',veh_type:'3T',vehicle_type_raw:'Chiller',weight:'3T',adhoc:'DM CARD / FOOD WATCH',partition_flag:'NO'},
  {vehicle_no:'AUH 17 68217',department:'Zooz',model:'MITSUBISHI CANTER',veh_type:'3T',vehicle_type_raw:'Chiller',weight:'3T',adhoc:'DM CARD / FOOD WATCH',partition_flag:'NO'},
  {vehicle_no:'AUH 14 17610',department:'Zooz',model:'MITSUBISHI CANTER',veh_type:'3T',vehicle_type_raw:'Chiller',weight:'3T',adhoc:'DM CARD / FOOD WATCH',partition_flag:'NO'},
  {vehicle_no:'AUH 20 91304',department:'Zooz',model:'MITSUBISHI CANTER',veh_type:'3T',vehicle_type_raw:'Chiller',weight:'3T',adhoc:'DM CARD / FOOD WATCH',partition_flag:'NO'},
  {vehicle_no:'AUH 8 54864',department:'Zooz',model:'TOYOTA HIACE',veh_type:'1T',vehicle_type_raw:'Chiller',weight:'1T',adhoc:'DM CARD / FOOD WATCH',partition_flag:'NO'},
  {vehicle_no:'AUH 14 26135',department:'Zooz',model:'TOYOTA HIACE',veh_type:'1T',vehicle_type_raw:'Chiller',weight:'1T',adhoc:'DM CARD / FOOD WATCH',partition_flag:'NO'},
  {vehicle_no:'AUH 14 25564',department:'Zooz',model:'TOYOTA HIACE',veh_type:'1T',vehicle_type_raw:'Chiller',weight:'1T',adhoc:'DM CARD / FOOD WATCH',partition_flag:'NO'},
  {vehicle_no:'AUH 8 54891',department:'Zooz',model:'TOYOTA HIACE',veh_type:'1T',vehicle_type_raw:'Chiller',weight:'1T',adhoc:'DM CARD / FOOD WATCH',partition_flag:'NO'},
  {vehicle_no:'AUH 13 30435',department:'Zooz',model:'TOYOTA HIACE',veh_type:'1T',vehicle_type_raw:'Chiller',weight:'1T',adhoc:'DM CARD / FOOD WATCH',partition_flag:'NO'},
  {vehicle_no:'DXB G 21566',department:'Zooz',model:'MITSUBISHI CANTER',veh_type:'3T',vehicle_type_raw:'chiller',weight:'3T',adhoc:'Dry Food',partition_flag:'n/a'},
  {vehicle_no:'DXB G 21790',department:'Zooz',model:'MITSUBISHI CANTER',veh_type:'3T',vehicle_type_raw:'Chiller',weight:'3T',adhoc:'DM CARD chiller',partition_flag:'n/a'},
  {vehicle_no:'DXB E 36231',department:'Zooz',model:'NISSAN URVAN',veh_type:'1T',vehicle_type_raw:'chiller',weight:'1T',adhoc:'-',partition_flag:'n/a'},
  {vehicle_no:'DXB J 74870',department:'Zooz',model:'TOYOTA HIACE',veh_type:'1T',vehicle_type_raw:'Chiller',weight:'1T',adhoc:'-',partition_flag:'n/a'},
  {vehicle_no:'DXB Y 79629',department:'Zooz',model:'TOYOTA HIACE',veh_type:'1T',vehicle_type_raw:'Chiller',weight:'1T',adhoc:'-',partition_flag:'n/a'},
  {vehicle_no:'DXB L 47349',department:'Zooz',model:'TOYOTA HIACE',veh_type:'1T',vehicle_type_raw:'Chiller',weight:'1T',adhoc:null,partition_flag:'n/a'},
  {vehicle_no:'DXB K 36814',department:'Zooz / Frozen',model:'Mitsubishi Canter',veh_type:'3T',vehicle_type_raw:'Frozen',weight:'-',adhoc:'DM CARD',partition_flag:'No'},
  {vehicle_no:'DXB K 36816',department:'Zooz / Frozen',model:'Mitsubishi Canter',veh_type:'3T',vehicle_type_raw:'Frozen',weight:'-',adhoc:'DM CARD',partition_flag:'No'},
  {vehicle_no:'DXB K 36817',department:'Zooz / Frozen',model:'Mitsubishi Canter',veh_type:'3T',vehicle_type_raw:'Frozen',weight:'-',adhoc:'DM CARD',partition_flag:'No'},
  {vehicle_no:'DXB K 30481',department:'Zooz / Frozen',model:'Mitsubishi Canter',veh_type:'3T',vehicle_type_raw:'Frozen',weight:'3T',adhoc:'DM CARD',partition_flag:'YES'},
  {vehicle_no:'AUH 8 55192',department:'Zooz / Frozen',model:'MITSUBISHI CANTER',veh_type:'3T',vehicle_type_raw:'Freezer',weight:'3T',adhoc:'DM CARD / FOOD WATCH',partition_flag:'YES'},
  {vehicle_no:'DXB K 36811',department:'Zooz / Frozen',model:'MITSUBISHI CANTER',veh_type:'3T',vehicle_type_raw:'Freezer',weight:'3T',adhoc:'DM CARD chiller',partition_flag:'n/a'}
];

module.exports = function (app, pool, requireAuth, requireRole, upload, auditLog, vehicleMasterMap) {

  function normalizeVehicleNo(raw) {
    return String(raw || '').toUpperCase().replace(/\s+/g, '');
  }

  async function seedIfEmpty() {
    try {
      const countRes = await pool.query(`SELECT COUNT(*)::int AS c FROM vehicle_master`);
      if (countRes.rows[0].c > 0) return; // already has data (either seeded before, or admin uploaded) — don't touch it
      let seeded = 0;
      for (const v of VEHICLE_MASTER_SEED_DATA) {
        const normalized = normalizeVehicleNo(v.vehicle_no);
        await pool.query(
          `INSERT INTO vehicle_master (vehicle_no, vehicle_no_normalized, department, model, veh_type, vehicle_type_raw, weight, adhoc, partition_flag)
           VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9) ON CONFLICT (vehicle_no_normalized) DO NOTHING`,
          [v.vehicle_no, normalized, v.department, v.model, v.veh_type, v.vehicle_type_raw, v.weight, v.adhoc, v.partition_flag]
        );
        seeded++;
      }
      console.log('Vehicle Master module: seeded ' + seeded + ' vehicles from embedded fleet list');
    } catch (e) {
      console.error('Vehicle Master seed error:', e.message);
    }
  }

  async function initVehicleMasterDB() {
    try {
      await pool.query(`CREATE TABLE IF NOT EXISTS vehicle_master (
        id SERIAL PRIMARY KEY,
        vehicle_no TEXT UNIQUE NOT NULL,
        vehicle_no_normalized TEXT UNIQUE NOT NULL,
        department TEXT,
        model TEXT,
        veh_type TEXT,
        vehicle_type_raw TEXT,
        weight TEXT,
        adhoc TEXT,
        partition_flag TEXT,
        active BOOLEAN DEFAULT true,
        created_at TIMESTAMPTZ DEFAULT NOW(),
        updated_at TIMESTAMPTZ DEFAULT NOW()
      )`);
      await pool.query(`CREATE INDEX IF NOT EXISTS idx_vehicle_master_norm ON vehicle_master(vehicle_no_normalized)`);
      console.log('Vehicle Master module: table ready');
      await seedIfEmpty();
      await refreshVehicleMasterMap();
    } catch (e) {
      console.error('Vehicle Master initDB error:', e.message);
    }
  }

  // Repopulate the shared in-memory map from the DB (mutates in place — never reassigns).
  // Indexed two ways: full normalized plate ("DXBN66510") AND trailing digits only ("66510") —
  // dispatch reports sometimes only carry the bare number without the city/prefix code.
  async function refreshVehicleMasterMap() {
    try {
      const r = await pool.query(`SELECT vehicle_no, vehicle_no_normalized, veh_type, vehicle_type_raw FROM vehicle_master WHERE active=true`);
      Object.keys(vehicleMasterMap).forEach(k => delete vehicleMasterMap[k]);
      r.rows.forEach(row => {
        const entry = { veh_type: row.veh_type, vehicle_type_raw: row.vehicle_type_raw };
        vehicleMasterMap[row.vehicle_no_normalized] = entry;
        // Extract the LAST separate digit group from the ORIGINAL vehicle number (before
        // spaces are stripped) — e.g. "AUH 18 18297" -> "18297", not "1818297".
        const digitGroups = String(row.vehicle_no || '').match(/\d+/g);
        if (digitGroups && digitGroups.length) {
          vehicleMasterMap['DIGITS:' + digitGroups[digitGroups.length - 1]] = entry;
        }
      });
      console.log('Vehicle Master module: ' + r.rows.length + ' vehicles loaded into memory');
    } catch (e) {
      console.error('Vehicle Master map refresh error:', e.message);
    }
  }
  initVehicleMasterDB();

  // Upload the vehicle master Excel (add/update only, never delete — same pattern as HoReCa masters)
  app.post('/api/vehicle-master/upload', requireAuth, requireRole('superadmin', 'subadmin'), upload.single('file'), async function (req, res) {
    try {
      if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
      const wb = XLSX.read(req.file.buffer, { type: 'buffer' });
      const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: null });

      let inserted = 0, updated = 0, skipped = 0;
      for (const row of rows) {
        const vehicleNo = String(row['Vehicle No.'] || row['Vehicle No'] || '').trim();
        if (!vehicleNo) { skipped++; continue; }
        const normalized = normalizeVehicleNo(vehicleNo);
        const department = row['Department'] || null;
        const model = row['Model'] || null;
        const vehType = row['Veh Type'] || null;
        const vehicleTypeRaw = row['Vehicle type'] || row['Vehicle Type'] || null;
        const weight = row['Weight'] || null;
        const adhoc = row['ADHOC'] || null;
        const partitionFlag = row['Partition'] || row['Partition '] || null;

        const existing = await pool.query(`SELECT id FROM vehicle_master WHERE vehicle_no_normalized=$1`, [normalized]);
        if (existing.rows.length) {
          await pool.query(
            `UPDATE vehicle_master SET vehicle_no=$1, department=$2, model=$3, veh_type=$4, vehicle_type_raw=$5, weight=$6, adhoc=$7, partition_flag=$8, updated_at=NOW()
             WHERE id=$9`,
            [vehicleNo, department, model, vehType, vehicleTypeRaw, weight, adhoc, partitionFlag, existing.rows[0].id]
          );
          updated++;
        } else {
          await pool.query(
            `INSERT INTO vehicle_master (vehicle_no, vehicle_no_normalized, department, model, veh_type, vehicle_type_raw, weight, adhoc, partition_flag)
             VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9)`,
            [vehicleNo, normalized, department, model, vehType, vehicleTypeRaw, weight, adhoc, partitionFlag]
          );
          inserted++;
        }
      }
      await refreshVehicleMasterMap();
      await auditLog(req.user.uid, req.user.username, 'VEHICLE_MASTER_UPLOAD', `inserted=${inserted} updated=${updated} skipped=${skipped}`, '');
      res.json({ success: true, inserted, updated, skipped, total_rows: rows.length });
    } catch (e) {
      res.status(500).json({ error: e.message });
    }
  });

  // Download what's currently stored, for verification
  app.get('/api/vehicle-master/export', requireAuth, requireRole('superadmin', 'subadmin'), async function (req, res) {
    try {
      const r = await pool.query(
        `SELECT vehicle_no AS "Vehicle No.", department AS "Department", model AS "Model", veh_type AS "Veh Type",
                vehicle_type_raw AS "Vehicle type", weight AS "Weight", adhoc AS "ADHOC", partition_flag AS "Partition"
         FROM vehicle_master WHERE active=true ORDER BY vehicle_no`
      );
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('Vehicle Master');
      if (r.rows.length) {
        ws.columns = Object.keys(r.rows[0]).map(k => ({ header: k, key: k, width: 18 }));
        r.rows.forEach(row => ws.addRow(row));
      }
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      res.setHeader('Content-Disposition', 'attachment; filename="Vehicle_Master_Current.xlsx"');
      await wb.xlsx.write(res);
      res.end();
    } catch (e) {
      if (!res.headersSent) res.status(500).json({ error: e.message });
    }
  });

  console.log('Vehicle Master module: routes mounted');
};
