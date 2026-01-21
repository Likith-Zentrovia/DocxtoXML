<!-- rittintermod.mod -->
<!-- This file contains a list of elements that are declared between the pool element set and the hierarchy -->
<!-- element set in the Rittenhose variation of the DocBook DTD -->
<!-- redeclarations are preformed in many ways -->
<!-- refentry  -->
<!ELEMENT risempty EMPTY>
<!ATTLIST risempty
		%common.attrib;
>

<!ENTITY % local.refentry.class "">
<!ENTITY % refentry.class	"risempty %local.refentry.class;">

<!-- local.usingsimplelist.class is defined in rittredef.mod -->
<!ENTITY % local.indexdivcomponent.mix "">
<!ENTITY % indexdivcomponent.mix
		"itemizedlist|orderedlist
		%local.usingvariablelist.class; %local.usingsimplelist.class;
		|%linespecific.class;	|%synop.class;
		|%para.class;		|%informal.class;
		|anchor %local.using.remark.class;
		|%link.char.class;
 		                        |beginpage
		%local.indexdivcomponent.mix;">

<!ENTITY % sect.title.content
	"title, subtitle?, titleabbrev?">

<!ENTITY % indexdiv.element "IGNORE">
<!ELEMENT indexdiv %ho; ((%sect.title.content;)?, ((%indexdivcomponent.mix;)*,
		(indexentry+  %local.usingsegmentedlist.class; )))>
<!--end of indexdiv.element-->

<!ENTITY % local.refname.char.mix "">
<!ENTITY % refname.char.mix
		"#PCDATA
		%local.barfor.tech.char.class; %tech.char.class;
		%local.refname.char.mix;">


<!-- bogusterm inclusion -->
<!ENTITY % bogusterm.module "IGNORE">
<![%bogusterm.module;[
<!ENTITY % local.bogusterm.attrib "">
<!ENTITY % bogusterm.role.attrib "%role.attrib;">

<!ENTITY % bogusterm.element "INCLUDE">
<![%bogusterm.element;[
<!ELEMENT bogusterm %ho; (#PCDATA)*>
<!--end of bogusterm.element-->]]>

<!ENTITY % bogusterm.attlist "INCLUDE">
<![%bogusterm.attlist;[
<!ATTLIST bogusterm
		%common.attrib;
		%bogusterm.role.attrib;
		%local.bogusterm.attrib;
>

<!ENTITY % local.refinline.char.mix "| bogusterm">
<!--end of bogusterm.attlist-->]]>
<!--end of bogusterm.module-->]]>

<!ENTITY % local.refinline.char.mix "">

<!ENTITY % refinline.char.mix
		"#PCDATA
		|%xref.char.class;	|%gen.char.class;
		|%link.char.class; %local.barfor.tech.char.class; %tech.char.class;
		|%base.char.class;	|%docinfo.char.class;
		|%other.char.class;
		|%ndxterm.class;        |beginpage
		%local.refinline.char.mix;">

