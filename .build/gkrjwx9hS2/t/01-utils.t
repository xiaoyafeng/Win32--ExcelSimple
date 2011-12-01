use strict;
use Data::Dumper;
use Test::More;
use Test::Deep;
use Test::Harness;
use Win32::ExcelSimple; 


is(cell2cr('AB3'),  (28, 3), "convert AB3 to [28, 3]");
is(cr2cell(2,34), 'B34', "convert (2,34) to B34");
done_testing;

