use strict;
use warnings;
use utf8;

use FindBin;
use Test::More;

use Data::XLSX::Parser;

my $parser = Data::XLSX::Parser->new;

my $fn = __FILE__;
$fn =~ s{t$}{xlsx};

$parser->open($fn);

my @sheets = $parser->workbook->names;
is scalar @sheets, 1, '1 worksheets ok';

is $sheets[0], 'sheet1', 'sheet1 name ok';

my $cells = [];
$parser->add_row_event_handler(sub {
    my ($row) = @_;
    push @$cells, $row;
});

$parser->sheet(1);

is $cells->[1][1], 1;
is $cells->[2][1], -3;
is $cells->[3][1], 1.5;
is $cells->[4][1], -1.5;
is $cells->[5][1], '3.5';
is $cells->[6][1], Time::Piece::gmtime(1400198400); # 2014/05/16 00:00:00
is $cells->[7][1], Time::Piece::gmtime(1400230800); # 2014/05/16 09:00:00
is $cells->[8][1], 'string';
is $cells->[9][1], '文字列';
is $cells->[10][1], 125702689;


done_testing;





