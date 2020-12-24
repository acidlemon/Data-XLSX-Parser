package Data::XLSX::Parser::Workbook;
use strict;
use warnings;

use XML::Parser::Expat;
use Archive::Zip ();
use File::Temp;
use Carp;

sub new {
    my ($class, $archive) = @_;

    my $self = bless [], $class;

    my $fh = File::Temp->new( SUFFIX => '.xml' ) or confess "couldn't create tempfile $!";
	
    my $handle = $archive->workbook or confess "couldn't get handle to workbook archive $!";;
    confess 'Failed to write temporally file: ', $fh->filename
        unless $handle->extractToFileNamed($fh->filename) == Archive::Zip::AZ_OK;

    my $parser = XML::Parser::Expat->new(Namespaces=>1);
    $parser->setHandlers(
        Start => sub { $self->_start(@_) },
        End   => sub {},
        Char  => sub {},
    );
    $parser->parse($fh);

    $self;
}

sub names {
    my ($self) = @_;
    map { $_->{name} } @$self;
}

sub sheet_rid {
    my ($self, $name) = @_;

    my ($meta) = grep { $_->{name} eq $name } @$self
        or return;

    return  $meta->{'id'};
}

sub sheet_id {
    my ($self, $name) = @_;

    my ($meta) = grep { $_->{name} eq $name } @$self
        or return;

     return $meta->{sheetId};
}

sub _start {
    my ($self, $parser, $el, %attr) = @_;
    push @$self, \%attr if $el eq 'sheet';
}

1;
