package Data::XLSX::Parser::DocumentArchive;
use strict;
use warnings;
use Carp;

use Archive::Zip;

sub new {
    my ($class, $filename) = @_;

    my $zip = Archive::Zip->new;
    if ($zip->read($filename) != Archive::Zip::AZ_OK) {
        confess "Cannot open file: $filename";
    }

    bless {
        _zip => $zip,
    }, $class;
}

sub workbook {
    my ($self) = @_;
    $self->{_zip}->memberNamed('xl/workbook.xml');
}

sub sheet {
    my ($self, $path) = @_;
	# only add xl/ if not already there, as some tools add absolute paths in relations
	$path = sprintf ('xl/%s', $path) unless $path =~ /^xl\//;
    $self->{_zip}->memberNamed($path);
}

sub shared_strings {
    my ($self) = @_;
    $self->{_zip}->memberNamed('xl/sharedStrings.xml');
}

sub styles {
    my ($self) = @_;
    $self->{_zip}->memberNamed('xl/styles.xml');
}

sub relationships {
    my ($self) = @_;
    $self->{_zip}->memberNamed('xl/_rels/workbook.xml.rels');
}

1;
