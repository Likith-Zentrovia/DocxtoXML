<!-- ritthier2.mod -->
<!-- This file contains a elements that are declared at the second hierarchy insertion point-->
<!-- element set in the Rittenhose variation of the DocBook DTD -->
<!-- redeclarations are preformed in many ways -->

<!-- | reference -->
<!ENTITY % div.title.content
	"title, subtitle?, titleabbrev?">

<!ENTITY % book.element "IGNORE">
<!ELEMENT book %ho; ((%div.title.content;)?, bookinfo?,
 		(dedication | toc | lot
 		| glossary | bibliography | preface
		| %chapter.class; | part 
		| %article.class;
 		| %appendix.class;
		| %index.class;
		| colophon)*)
		%ubiq.inclusion;>

<!-- entity that is used befor declariont -->
<!ENTITY % bookcomponent.title.content
	"title, subtitle?, titleabbrev?">
	
<!ENTITY % chapter.element "IGNORE">
<!-- redefinition of chapter to force sect1 tagging -->
<!-- removed (%bookcomponent.content;),-->
<!ENTITY % using.bookcomponent.content "">
<!ENTITY % using.sect1only.content " | sect1">
<!--  using.sect1only.content  -->
<!-- Fixed non-deterministic content model by removing (%nav.class;)* before tocchap -->
<!ELEMENT chapter %ho; (beginpage?,
                    chapterinfo?,
                    (%bookcomponent.title.content;),
                    tocchap?,
                    %using.bookcomponent.content;
                    ( %nav.class; %using.sect1only.content;)*)
		%ubiq.inclusion;>
<!--end of chapter.element-->
<!ENTITY 	% local.sect1.attrib "" >
<!ENTITY 	% local.sect2.attrib "" >
<!ENTITY 	% local.sect3.attrib "" >
<!ENTITY 	% local.sect4.attrib "" >
<!ENTITY 	% local.sect5.attrib "" >
<!ENTITY % sect1.role.attrib "%role.attrib;">
<!ENTITY % sect2.role.attrib "%role.attrib;">
<!ENTITY % sect3.role.attrib "%role.attrib;">
<!ENTITY % sect4.role.attrib "%role.attrib;">
<!ENTITY % sect5.role.attrib "%role.attrib;">

<!-- attribute list redefintion for early sections. -->
<!ENTITY % sect1.attlist "IGNORE">
<!ATTLIST sect1
		renderas	(sect2
				|sect3
				|sect4
				|sect5)		#IMPLIED
		%label.attrib;
		%status.attrib;
		%idreq.common.attrib;
		%sect1.role.attrib;
		%local.sect1.attrib;
>
<!--end of sect1.attlist-->

<!ENTITY % sect2.attlist "IGNORE">
<!ATTLIST sect2
		renderas	(sect1
				|sect3
				|sect4
				|sect5)		#IMPLIED
		%label.attrib;
		%status.attrib;
		%idreq.common.attrib;
		%sect2.role.attrib;
		%local.sect2.attrib;
>
<!--end of sect2.attlist-->

<!ENTITY % sect3.attlist "IGNORE">
<!ATTLIST sect3
		renderas	(sect1
				|sect2
				|sect4
				|sect5)		#IMPLIED
		%label.attrib;
		%status.attrib;
		%idreq.common.attrib;
		%sect3.role.attrib;
		%local.sect3.attrib;
>
<!--end of sect3.attlist-->

<!ENTITY % sect4.attlist "IGNORE">
<!ATTLIST sect4
		renderas	(sect1
				|sect2
				|sect3
				|sect5)		#IMPLIED
		%label.attrib;
		%status.attrib;
		%idreq.common.attrib;
		%sect4.role.attrib;
		%local.sect4.attrib;
>
<!--end of sect4.attlist-->

<!ENTITY % sect5.attlist "IGNORE">
<!ATTLIST sect5
		renderas	(sect1
				|sect2
				|sect3
				|sect4)		#IMPLIED
		%label.attrib;
		%status.attrib;
		%idreq.common.attrib;
		%sect5.role.attrib;
		%local.sect5.attrib;
>
<!--end of sect5.attlist-->

<!-- check for inclusion of custom tags -->
<![%rittcustomtags.redecl.module;[
%rittcustomtags;
]]>

<!-- Add sect6 to content model of sect5 -->
<!ENTITY % sect5.role.attrib "%role.attrib;">
<!ELEMENT sect5 %ho; (sect5info?, (%sect.title.content;), (%nav.class;)*,
        (((%divcomponent.mix;)+, 
                ((%refentry.class;)* | sect6* | simplesect*))
        | (%refentry.class;)+ | sect6+ | simplesect+), (%nav.class;)*)>

<!-- add 6-10 info blocks -->
<!ENTITY % local.sect6info.attrib "">
<!ENTITY % sect6info.role.attrib "%role.attrib;">

<!ENTITY % sect6info.element "INCLUDE">
<![%sect6info.element;[
<!ELEMENT sect6info %ho; ((%info.class;)+)
		%beginpage.exclusion;>
<!--end of sect6info.element-->]]>

<!ENTITY % sect6info.attlist "INCLUDE">
<![%sect6info.attlist;[
<!ATTLIST sect6info
		%common.attrib;
		%sect6info.role.attrib;
		%local.sect6info.attrib;
>
<!--end of sect6info.attlist-->]]>

<!ENTITY % local.sect7info.attrib "">
<!ENTITY % sect7info.role.attrib "%role.attrib;">

<!ENTITY % sect7info.element "INCLUDE">
<![%sect7info.element;[
<!ELEMENT sect7info %ho; ((%info.class;)+)
		%beginpage.exclusion;>
<!--end of sect7info.element-->]]>

<!ENTITY % sect7info.attlist "INCLUDE">
<![%sect7info.attlist;[
<!ATTLIST sect7info
		%common.attrib;
		%sect7info.role.attrib;
		%local.sect7info.attrib;
>
<!--end of sect7info.attlist-->]]>

<!ENTITY % local.sect8info.attrib "">
<!ENTITY % sect8info.role.attrib "%role.attrib;">

<!ENTITY % sect8info.element "INCLUDE">
<![%sect8info.element;[
<!ELEMENT sect8info %ho; ((%info.class;)+)
		%beginpage.exclusion;>
<!--end of sect8info.element-->]]>

<!ENTITY % sect8info.attlist "INCLUDE">
<![%sect8info.attlist;[
<!ATTLIST sect8info
		%common.attrib;
		%sect8info.role.attrib;
		%local.sect8info.attrib;
>
<!--end of sect8info.attlist-->]]>

<!ENTITY % local.sect9info.attrib "">
<!ENTITY % sect9info.role.attrib "%role.attrib;">

<!ENTITY % sect9info.element "INCLUDE">
<![%sect9info.element;[
<!ELEMENT sect9info %ho; ((%info.class;)+)
		%beginpage.exclusion;>
<!--end of sect9info.element-->]]>

<!ENTITY % sect9info.attlist "INCLUDE">
<![%sect9info.attlist;[
<!ATTLIST sect9info
		%common.attrib;
		%sect9info.role.attrib;
		%local.sect9info.attrib;
>
<!--end of sect9info.attlist-->]]>

<!ENTITY % local.sect10info.attrib "">
<!ENTITY % sect10info.role.attrib "%role.attrib;">

<!ENTITY % sect10info.element "INCLUDE">
<![%sect10info.element;[
<!ELEMENT sect10info %ho; ((%info.class;)+)
		%beginpage.exclusion;>
<!--end of sect10info.element-->]]>

<!ENTITY % sect10info.attlist "INCLUDE">
<![%sect10info.attlist;[
<!ATTLIST sect10info
		%common.attrib;
		%sect10info.role.attrib;
		%local.sect10info.attrib;
>
<!--end of sect10info.attlist-->]]>




<!-- end info blocks -->

<!-- and 6-10 element blocks -->
<!-- sect6 -->
<!ENTITY % sect6.role.attrib "%role.attrib;">
<!ELEMENT sect6 %ho; (sect6info?, (%sect.title.content;), (%nav.class;)*,
        (((%divcomponent.mix;)+, ((%refentry.class;)* | sect7* | simplesect*))
        | (%refentry.class;)+ | sect7+ | simplesect+), (%nav.class;)*)>
<!ATTLIST sect6
		renderas	(sect1
				|sect2
				|sect3
				|sect4
				|sect5
				|sect7
				|sect8
				|sect9
				|sect10)		#IMPLIED
        %label.attrib;
        %status.attrib;
        %idreq.common.attrib;
        %sect6.role.attrib;
>

<!-- sect7 -->
<!ENTITY % sect7.role.attrib "%role.attrib;">
<!ELEMENT sect7 %ho; (sect7info?, (%sect.title.content;), (%nav.class;)*,
        (((%divcomponent.mix;)+, ((%refentry.class;)* | sect8* | simplesect*))
        | (%refentry.class;)+ | sect8+  | simplesect+), (%nav.class;)*)>
<!ATTLIST sect7
		renderas	(sect1
				|sect2
				|sect3
				|sect4
				|sect5
				|sect6
				|sect8
				|sect9
				|sect10)		#IMPLIED
        %label.attrib;
        %status.attrib;
        %idreq.common.attrib;
        %sect7.role.attrib;
>
<!-- sect8 -->
<!ENTITY % sect8.role.attrib "%role.attrib;">
<!ELEMENT sect8 %ho; (sect8info?, (%sect.title.content;), (%nav.class;)*,
        (((%divcomponent.mix;)+, ((%refentry.class;)* | sect9* | simplesect*))
        | (%refentry.class;)+ | sect9+ | simplesect+), (%nav.class;)*)>
<!ATTLIST sect8
		renderas	(sect1
				|sect2
				|sect3
				|sect4
				|sect5
				|sect6
				|sect7
				|sect9
				|sect10)		#IMPLIED
        %label.attrib;
        %status.attrib;
        %idreq.common.attrib;
        %sect8.role.attrib;
>
<!-- sect9 -->
<!ENTITY % sect9.role.attrib "%role.attrib;">
<!ELEMENT sect9 %ho; (sect9info?, (%sect.title.content;), (%nav.class;)*,
        (((%divcomponent.mix;)+, ((%refentry.class;)* | sect10* | simplesect*))
        | (%refentry.class;)+ | sect10+ | simplesect+), (%nav.class;)*)>
<!ATTLIST sect9
		renderas	(sect1
				|sect2
				|sect3
				|sect4
				|sect5
				|sect6
				|sect7
				|sect8
				|sect10)		#IMPLIED
        %label.attrib;
        %status.attrib;
        %idreq.common.attrib;
        %sect9.role.attrib;
>
<!-- sect10 -->
<!ENTITY % sect10.role.attrib "%role.attrib;">
<!ELEMENT sect10 %ho; (sect10info?, (%sect.title.content;), (%nav.class;)*,
        (((%divcomponent.mix;)+, ((%refentry.class;)* | simplesect*))
        | (%refentry.class;)+ | simplesect+), (%nav.class;)*)>
<!ATTLIST sect10
		renderas	(sect1
				|sect2
				|sect3
				|sect4
				|sect5
				|sect6
				|sect7
				|sect8
				|sect9)		#IMPLIED
        %label.attrib;
        %status.attrib;
        %idreq.common.attrib;
        %sect10.role.attrib;
>
