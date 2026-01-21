<!-- rittredef.mod-->
<!-- This file contains a list of elements that are redefined in the the Rittenhouse variation of the DocBook DTD -->
<!-- redefinitions are preformed in various ways -->
<!--  procedure in mix -->

<!ENTITY % local.sidebar.mix "">
<!ENTITY % sidebar.mix
		"%list.class;		|%admon.class;
		|%linespecific.class;	|%synop.class;
		|%para.class;		|%informal.class;
		|%formal.class;		%local.usingprocedure.class;
		|%genobj.class;
		|%ndxterm.class;        |beginpage
		%local.sidebar.mix;">

<!ENTITY % local.qandaset.mix "">
<!ENTITY % qandaset.mix
		"%list.class;           |%admon.class;
		|%linespecific.class;	|%synop.class;
		|%para.class;		|%informal.class;
		|%formal.class;		%local.usingprocedure.class;
		|%genobj.class;
		|%ndxterm.class;
		%local.qandaset.mix;">

<!ENTITY % local.revdescription.mix "">
<!ENTITY % revdescription.mix
		"%list.class;		|%admon.class;
		|%linespecific.class;	|%synop.class;
		|%para.class;		|%informal.class;
		|%formal.class;		%local.usingprocedure.class;
		|%genobj.class;
		|%ndxterm.class;
		%local.revdescription.mix;">
		
<!ENTITY % local.using.remark.class "">
<!ENTITY % local.admon.mix "">
<!ENTITY % admon.mix
		"%list.class;
		|%linespecific.class;	|%synop.class;
		|%para.class;		|%informal.class;
		|%formal.class;		%local.usingprocedure.class;
		|anchor|bridgehead %local.using.remark.class;
		|%ndxterm.class;        |beginpage
		%local.admon.mix;">

<!ENTITY % local.textobject.mix "">
<!ENTITY % textobject.mix
		"%list.class;		|%admon.class;
		|%linespecific.class;
		|%para.class;		|blockquote
		%local.textobject.mix;">

<!-- REMOVAL OF CONF GROUP -->
<!-- person must beredefined for ordering -->
<!ENTITY % local.person.ident.mix "">
<!ENTITY % person.ident.mix
		"honorific|firstname|surname|lineage|othername|affiliation
		|authorblurb|contrib %local.person.ident.mix;">

<!ENTITY % local.bibliocomponent.mix "">

<!ENTITY % local.usingconfgroup.class "">
<!ENTITY % local.using.citebiblioid.class "">
<!ENTITY % local.usingcontractsponsor.class "">
<!ENTITY % bibliocomponent.mix
		"abbrev %local.usingabstract.class; |address|artpagenums|author
		|authorgroup|authorinitials|bibliomisc|biblioset
		|collab %local.usingconfgroup.class;|contractnum %local.usingcontractsponsor.class;
		|copyright|corpauthor|corpname|date|edition
		|editor|invpartnumber|isbn|issn|issuenum|orgname
		|biblioid %local.using.citebiblioid.class;
		|bibliosource|bibliorelation|bibliocoverage
		|othercredit|pagenums|printhistory|productname
		|productnumber|pubdate|publisher|publishername
		|pubsnumber|releaseinfo|revhistory|seriesvolnums
		|subtitle|title|titleabbrev|volumenum|citetitle
		|personname|%person.ident.mix;
		|%ndxterm.class;
		%local.bibliocomponent.mix;">

<!ENTITY % local.refclass.char.mix "">
<!--- application is "|application" if used -->
<!ENTITY % local.usingapplication.class "">
<!ENTITY % refclass.char.mix
		"#PCDATA
		%local.usingapplication.class;
		%local.refclass.char.mix;">
		
<!ENTITY % local.usinginformalequation.class "">

<!ENTITY % equation.module "INCLUDE">
<![%equation.module;[
<!ENTITY % equation.element "IGNORE">
<!ELEMENT equation %ho; (blockinfo?, (%formalobject.title.content;)?,
                         ( %local.usinginformalequation.class; %equation.content;))>
<!--end of equation.module -->]]>

<!ENTITY % forminlines.hook "">

<!-- I don't think this is well placed, but it needs to be here because of -->
<!-- the reference to bibliocomponent.mix -->
<!-- keyword exclusion -->
<!ENTITY % local.usingkeywordset.class "| keywordset">

<!ENTITY % local.info.class "|primaryauthor|risinfo">
<!ENTITY % info.class
		"graphic | mediaobject | legalnotice %local.usingmodespec.class;
		 | subjectset   %local.usingkeywordset.class;
		 | itermset | %bibliocomponent.mix;
                 %local.info.class;">

<!ENTITY % ubiq.exclusion "">

<!ENTITY % local.para.char.mix "">
<!ENTITY % para.char.mix
		"#PCDATA
		|%xref.char.class;	|%gen.char.class;
		|%link.char.class;	%local.barfor.tech.char.class;   %tech.char.class;
		|%base.char.class;	|%docinfo.char.class;
		|%other.char.class;	|%inlineobj.char.class;
		|%synop.class;
		|%ndxterm.class;        |beginpage
                %forminlines.hook;
		%local.para.char.mix;">


<!ENTITY % local.title.char.mix "">
<!ENTITY % title.char.mix
		"#PCDATA
		|%xref.char.class;	|%gen.char.class;
		|%link.char.class;	%local.barfor.tech.char.class;  %tech.char.class;
		|%base.char.class;	|%docinfo.char.class;
		|%other.char.class;	|%inlineobj.char.class;
		|%ndxterm.class;
		%local.title.char.mix;">


<!ENTITY % local.ndxterm.char.mix "">
<!ENTITY % ndxterm.char.mix
		"#PCDATA
		|%xref.char.class;	|%gen.char.class;
		|%link.char.class;	%local.barfor.tech.char.class;  %tech.char.class;
		|%base.char.class;	|%docinfo.char.class;
		|%other.char.class;	|inlinegraphic|inlinemediaobject
		%local.ndxterm.char.mix;">

<!ENTITY % local.cptr.char.mix "">
<!ENTITY % cptr.char.mix
		"#PCDATA
		|%link.char.class;	%local.barfor.tech.char.class; %tech.char.class;
		|%base.char.class;
		|%other.char.class;	|inlinegraphic|inlinemediaobject
		|%ndxterm.class;        |beginpage
		%local.cptr.char.mix;">


<!ENTITY % local.usingreplaceable.class "">

<!ENTITY % local.smallcptr.char.mix "">
<!ENTITY % smallcptr.char.mix
		"#PCDATA
					%local.usingreplaceable.class;
					|inlinegraphic|inlinemediaobject
		|%ndxterm.class;        |beginpage
		%local.smallcptr.char.mix;">

<!ENTITY % local.docinfo.char.mix "">
<!ENTITY % docinfo.char.mix
		"#PCDATA
		|%link.char.class;
					|emphasis|trademark
					%local.usingreplaceable.class;
		|%other.char.class;	|inlinegraphic|inlinemediaobject
		|%ndxterm.class;
		%local.docinfo.char.mix;">

<!ENTITY % subscript.element "IGNORE">
<!ELEMENT subscript %ho; (#PCDATA
		| %link.char.class;
		| emphasis
		%local.usingreplaceable.class;
		%local.using.symbol.class;
		| inlinegraphic
                | inlinemediaobject
		| %base.char.class;
		| %other.char.class;)*
		%ubiq.exclusion;>

<!ENTITY % superscript.element "IGNORE">

<!ELEMENT superscript %ho; (#PCDATA
		| %link.char.class;
		| emphasis
		%local.usingreplaceable.class;
		%local.using.symbol.class;
		| inlinegraphic
                | inlinemediaobject
		| %base.char.class;
		| %other.char.class;)*
		%ubiq.exclusion;>


<!ENTITY % local.usingaccel.class "">
<!ENTITY % interface.element "IGNORE">

<!ELEMENT interface %ho; (%smallcptr.char.mix; %local.usingaccel.class; )*>


<!ENTITY % local.para.char.mix "">
<!ENTITY % para.char.mix
		"#PCDATA
		|%xref.char.class;	|%gen.char.class;
		|%link.char.class;	%local.barfor.tech.char.class; %tech.char.class;
		|%base.char.class;	|%docinfo.char.class;
		|%other.char.class;	|%inlineobj.char.class;
		|%synop.class;
		|%ndxterm.class;        |beginpage
                %forminlines.hook;
		%local.para.char.mix;">

<!ENTITY % local.usinglineannotation.class "">
<!ENTITY % local.using.co.class "">
<!-- |co|coref -->
<!ENTITY % literallayout.element "IGNORE">
<!ELEMENT literallayout %ho; 
	(%para.char.mix;%local.using.co.class;|textobject 
	%local.usinglineannotation.class;)*>
<!--end of literallayout.element-->

<!ENTITY % synopsis.element "IGNORE">
<!ELEMENT synopsis %ho; (
	%para.char.mix;|graphic|mediaobject%local.using.co.class;|textobject 
	%local.usinglineannotation.class; )*>
<!--end of synopsis.element-->

<!ENTITY % trademark.element "IGNORE">
<!ELEMENT trademark %ho; (#PCDATA
		| %link.char.class;
		%local.barfor.tech.char.class;  %tech.char.class;
		| %base.char.class;
		| %other.char.class;
		| inlinegraphic
                | inlinemediaobject
		| emphasis)*>
<!--end of trademark.element-->
<!-- |textdata -->
<!ENTITY % textobject.element "IGNORE">
<!ELEMENT textobject %ho; (objectinfo?, (phrase|(%textobject.mix;)+))>
<!--end of textobject.element-->

<!ENTITY % local.legalnotice.mix "">

