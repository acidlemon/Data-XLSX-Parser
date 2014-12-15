# NAME

Data::XLSX::Parser - faster XLSX parser

# SYNOPSIS

    use Data::Dumper;
    use Data::XLSX::Parser;
    
    my $parser = Data::XLSX::Parser->new;
    $parser->add_row_event_handler(sub {
        my ($row) = @_;
        print Dumper $row;
    });
    $parser->parse('foo.xlsx');

# DESCRIPTION

Data::XLSX::Parser provides faster way to parse Microsoft Excel's .xlsx files.
The implementation of this module is highly inspired from Python's FastXLSX library.

This is SAX based parser, so you can parse very large XLSX file with lower memory usage.

# THIS MODULE IS \*ALPHA\* QUALITY

This module is created for my current daily work that needs convert very huge excel file to csv, and perfectly work against my files but might not to all excel datas.

If you have some XSLX files that doesn't parse this module, please bug me with the files.

# METHODS

## new

Create new parser object.

## add

# AUTHOR

Daisuke Murase <typester@cpan.org>
