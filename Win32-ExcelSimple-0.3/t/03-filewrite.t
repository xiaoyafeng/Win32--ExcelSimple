use strict;
use Test::More;
use Test::Deep;
use Test::Harness;
use Win32::ExcelSimple;
use File::Basename;
use File::Spec::Functions;

my $path = dirname($0);
   $path = Win32::GetFullPathName($path);
my $abs_file = catfile($path, 'test.xls');
my $es = Win32::ExcelSimple->new($abs_file);
my $sheet_h = $es->open_sheet('Report');
   $sheet_h->write_cell(2,1, 'test');
is($sheet_h->read_cell(2,1),  'test', "write string test to cell B1");
   $sheet_h->write(2,1, ['test1','test2','test3']);
is_deeply($sheet_h->read(2,1,4,1),  ['test1','test2','test3'], "write data to a Range");
   $sheet_h->write(2,1, [
		                ['test1','test2','test3'],
						['chop1','chop2','chop3'],
						['deer1','deer2','deer3'],
					    ]
					);
cmp_deeply($sheet_h->read(2,1,4,3), [
		                ['test1','test2','test3'],
						['chop1','chop2','chop3'],
						['deer1','deer2','deer3'],
					    ],
					    "write data to a multi-dimension Range");  	
done_testing;

