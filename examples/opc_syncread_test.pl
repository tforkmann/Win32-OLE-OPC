use Win32::OLE::OPC;

# Karl Scherer <scherer.karl@gene.com> wrote this script to test the code he
# contributed to OPC.pm.

my $yoko   = $ARGV[0] ? $ARGV[0] : "SET27.YOKO051";
my $num    = $ARGV[1] ? $ARGV[1] : 3;
my $server = $ARGV[2] ? $ARGV[2] : "CIO-YOKO2";

my $debug = 0;

if ($debug) {
  print "\nARGS: $yoko  $num  $server\n";
}

my $opcintf = Win32::OLE::OPC->new(#'OCSTK.DAAutomation',
                                   'OPC.Automation',
                                   #'RSLinx.RSLinx',
                                   'KEPware.KEPServerEx.V4',
                                   $server);

%serverprop = $opcintf->ServerProperties;

if ($debug) {
  print scalar(keys(%serverprop)), " Server Properties\n";
  foreach $prop (keys %serverprop) {
    print "$prop  $serverprop{$prop}\n";
  }
}

$group = $opcintf->OPCGroups->Add('kks');

if ($debug) {
  %groupprop = $group->Properties;
  print scalar(keys(%groupprop)), " Group Properties\n";
  foreach $prop (keys %groupprop) {
    print "$prop  $groupprop{$prop}\n";
  }
}

$items = $group->OPCItems();

# put your OPC item names here
my $names = [qw(SET27.YOKO051.CH001
                SET27.YOKO051.CH001_Alarm
                SET27.YOKO051.CH001_status)
            ];
my $client_handles = [1,2,3];

# Very site/device specific. tested with 60 channels ~ 180 OPC items.
for ($i=1; $i<=$num; $i++) {
  $$names[$i-1] = $yoko.".CH".substr("000".$i,-3,3);
  $$names[$i-1+$num] = $yoko.".CH".substr("000".$i,-3,3)."_status";
  $$names[$i-1+2*$num] = $yoko.".CH".substr("000".$i,-3,3)."_Alarm";
  $$client_handles[$i-1] = $i;
  $$client_handles[$i-1+$num] = $i+$num;
  $$client_handles[$i-1+2*$num] = $i+2*$num;
}

my $multiple = 3;
$num *= $multiple ;

my $server_handles = [];
my $values = [];
my $errors = [];
my $quals  = [];
my $time_stamps = [];

if ($debug) {
  print "before AddItems:\n";
  for ($i=0; $i<$num; $i++) {
    print "$i $$names[$i] $$client_handles[$i] $$server_handles[$i] $$values[$i] $$errors[$i] $$quals[$i] $$time_stamps[$i]\n";
  }
}

$items->AddItems($num, $names, $client_handles, $errors, $server_handles);

if ($debug) {
  print "after AddItems, before SyncRead:\n";
  for ($i=0; $i<$num; $i++) {
    print "$i $$names[$i] $$server_handles[$i] $$values[$i] $$errors[$i] $$quals[$i] $$time_stamps[$i]\n";
  }
}

$group->SyncRead($OPCDevice, $num, $server_handles, $values, $errors, $quals, $time_stamps, $num);

if ($debug) {
  print "after SyncRead:\n";
  for ($i=0; $i<$num; $i++) {
    print "$i $$names[$i] $$server_handles[$i] $$values[$i] $$errors[$i] $$quals[$i] $$time_stamps[$i]\n";
  }
}

if (defined $multiple) {
  $num /= $multiple;
}
for ($i=0; $i<$num; $i++) {
  print "$$names[$i]\t$$values[$i]\t$$values[$i+$num]\t$$values[$i+2*$num]\t\[$$quals[$i] $$time_stamps[$i]\]\n";
}
print "finished\n";

exit;

