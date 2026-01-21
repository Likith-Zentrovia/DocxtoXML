<!-- rittredef.mod-->
<!-- This file contains a list of elements that are redefined in the the Rittenhouse variation of the DocBook DTD -->
<!-- redefinitions are preformed in various ways -->
<!-- msgset  and procedure-->

<!ENTITY % local.ndxterm.class "">
<!ENTITY % ndxterm.class
		"indexterm %local.ndxterm.class;">


<!ENTITY % ebnf.block.hook "">
<!-- placeholder for msgset -->
<!ENTITY % local.usingmsgset.class "">
<!-- placeholder for procedure -->
<!ENTITY % local.usingprocedure.class "">

<!-- placeholder for confgroup -->
<!ENTITY % local.usingconfgroup.class "">

<!-- oo holder value "|ooclass|oointerface|ooexception" -->
<!ENTITY % local.usingoo.class "">

<!-- placeholder for confgroup -->
<!ENTITY % local.usingclassname.class "">

<!ENTITY % local.compound.class "">
<!ENTITY % compound.class
		"sidebar|qandaset
                 %ebnf.block.hook;
                 %local.compound.class;
                 %local.usingmsgset.class;
                 %local.usingprocedure.class;
                 ">

<!ENTITY % local.using.remark.class "">
<!ENTITY % local.genobj.class "">
<!ENTITY % genobj.class
		"anchor|bridgehead %local.using.remark.class; |highlights
		%local.genobj.class;">

<!ENTITY % local.using.citerefentry.class "">

<!ENTITY % local.gen.char.class "">
<!ENTITY % gen.char.class
		"abbrev|acronym|citation %local.using.citerefentry.class;
		|citetitle|emphasis
		|firstterm|foreignphrase|glossterm|footnote|phrase|orgname
		|quote|trademark|wordasword|personname %local.gen.char.class;">
<!-- -->
<!ENTITY % local.usingonlink.class "">
<!ENTITY % local.link.char.class "">
<!ENTITY % link.char.class
		"link %local.usingonlink.class; |ulink %local.link.char.class;">

<!-- placeholder for abbstract-->
<!ENTITY % local.usingabstract.class "|abstract">

<!ENTITY % local.descobj.class "">
<!ENTITY % descobj.class
		"authorblurb|epigraph %local.usingabstract.class;
		%local.descobj.class;">

<!-- placeholder for alt-->
<!-- not defin would be alt?, -->
<!ENTITY % local.usingalt.class "">

<!ENTITY % equation.content "( %local.usingalt.class;(graphic+|mediaobject+))">
<!--<!ENTITY % inlineequation.content "(%local.usingalt.class; (graphic+|inlinemediaobject+))"> -->

<!ENTITY % local.tech.char.class "">
<!ENTITY % local.usingapplication.class "">
<!ENTITY % local.usingexceptionname.class "">
<!ENTITY % local.usingconstant.class "">
<!ENTITY % local.usingenvar.class "">
<!ENTITY % local.usingerror.class "">
<!-- errorclass includes |errorcode|errorname|errortype|errortext" -->

<!ENTITY % ebnf.inline.hook "">

<!-- function class -->
<!ENTITY % local.usingfunction.class "">
<!-- key class -->
<!ENTITY % local.usingkey.class "">
<!-- key class includes |errorcode|errorname|errortype|errortext" -->
<!-- gui class -->
<!ENTITY % local.usinggui.class "">
<!-- gui class includes |guibutton|guiicon|guilabel|guimenu|guimenuitem	|guisubmenu" -->
<!ENTITY % local.usinghardware.class "">
<!-- computer turems --> 
<!ENTITY % local.usinginterfaces.class "">
<!-- |methodname|interfacename -->
<!-- markup -->
<!ENTITY % local.usingmarkup.class "">

<!-- more interface -->
<!ENTITY % local.usingmoreinf1.class "">
<!-- |medialabel|menuchoice|mousebutton|option|optional | parameter|prompt|property|replaceable|returnvalue|sgmltag|structfield|structname -->
<!ENTITY % local.usingmoreinf2.class "">
<!-- |token|command|computeroutput| systemitem|userinput|varname -->

<!ENTITY % local.using.symbol.class "">

<!-- the bar "|" is inclued outside tech.char.class which causes the parser to fail if %barfor.tech.char.class; is redefined away which we have done -->
<!-- so the bar must be defined and the using mixes must be redfined -->
<!ENTITY % local.barfor.tech.char.class "">
<!ENTITY % tech.char.class
		"%local.usingapplication.class;
                %local.usingclassname.class; %local.usinginterfaces.class; %local.usingexceptionname.class;
                %local.usingoo.class;
                
		|database|email   %local.usingenvar.class; %local.usingerror.class;
		|filename
		%local.usingfunction.class;
		%local.usinggui.class;
		 %local.usinghardware.class;  |interface
		%local.usingkey.class;
		|literal %local.usingconstant.class; %local.usingmarkup.class; 
		%local.usingmoreinf1.class;
		%local.using.symbol.class; |type 
		%local.usingmoreinf2.class;
                %ebnf.inline.hook;
		%local.tech.char.class;">


<!ENTITY % local.other.char.class "">
<!ENTITY % other.char.class
		"subscript|superscript %local.using.remark.class; 
		%local.other.char.class;">

<!ENTITY % local.list.class "">
<!ENTITY % local.usingcalloutlist.class "">
<!ENTITY % local.usingsimplelist.class "">
<!ENTITY % local.usingvariablelist.class "">
<!ENTITY % local.usingsegmentedlist.class "">

<!ENTITY % list.class
		"%local.usingcalloutlist.class; glosslist|itemizedlist|orderedlist
		%local.usingsegmentedlist.class;
		 %local.usingsimplelist.class; %local.usingvariablelist.class; %local.list.class;">

<!-- screen removal -->
<!ENTITY % local.usingscreens.class "">
<!-- |screen|screenco|screenshot  -->
<!ENTITY % local.usingprogramlisting.class "">
<!-- |programlisting|programlistingco -->

<!ENTITY % local.linespecific.class "">
<!ENTITY % linespecific.class
		"literallayout   %local.usingprogramlisting.class; 
		%local.usingscreens.class; %local.linespecific.class;">

<!-- paragraph redef -->
<!ENTITY % local.para.class "">
<!ENTITY % local.usingsimpara.class "">
<!ENTITY % para.class
		"formalpara |para %local.usingsimpara.class; %local.para.class;">


<!-- graphic co removal -->
<!ENTITY % local.usinggraphicco.class "">

<!ENTITY % local.usinginformal.class "">
<!-- |informalequation |informalexample |informalfigure |informaltable  -->
<!-- mediaobjectco -->
<!ENTITY % local.usingmediaobjectco.class "">

<!ENTITY % local.informal.class "">
<!ENTITY % informal.class
		"address|blockquote
                |graphic %local.usinggraphicco.class;  |mediaobject %local.usingmediaobjectco.class;
                %local.usinginformal.class;
                %local.informal.class;">


<!-- removed bunch techical synopsis types-->
<!-- |cmdsynopsis|funcsynopsis |classsynopsis|fieldsynopsis |%method.synop.class;-->

<!ENTITY % local.usingtectsynop.class "">

<!ENTITY % local.synop.class "">
<!ENTITY % synop.class
		"synopsis %local.usingtectsynop.class;
                  %local.synop.class;">

<!ENTITY % local.usingexample.class "">
<!ENTITY % local.formal.class "">
<!ENTITY % formal.class
		"equation  %local.usingexample.class;   |figure|table %local.formal.class;">

<!-- remove inlineexquation 
<!ENTITY % local.usinginlineequation.class "">
<!ENTITY % local.inlineobj.char.class "">
<!ENTITY % inlineobj.char.class
		"inlinegraphic|inlinemediaobject %local.usinginlineequation.class;  %local.inlineobj.char.class;">
-->
<!ENTITY % local.usingmodespec.class "">
<!ENTITY % local.docinfo.char.class "">
<!ENTITY % docinfo.char.class
		"author|authorinitials|corpauthor %local.usingmodespec.class; |othercredit
		|productname|productnumber|revhistory
		%local.docinfo.char.class;">

<!ENTITY % local.glossdef.mix "">
<!ENTITY % glossdef.mix
		"%list.class;
		|%linespecific.class;	|%synop.class;
		|%para.class;		|%informal.class;
		|%formal.class;
		%local.using.remark.class;
		|%ndxterm.class;        |beginpage
		%local.glossdef.mix;">


<!-- degree element for author degrees -->

<!ENTITY % local.personname.element "INCLUDE">
<![%local.personname.element;[
<!ENTITY % personname.element "IGNORE">
<!ELEMENT degree		(#PCDATA) >
<!-- added degree-->
<!ELEMENT personname %ho; ((honorific|firstname|surname|lineage|othername|degree)+)>
<!ENTITY % local.person.ident.mix "|degree">
<!--end of personname.element-->]]>
