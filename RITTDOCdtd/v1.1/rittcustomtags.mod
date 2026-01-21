<!-- rittcustomtags.mod -->
<!-- This file contains modifications that are local to the rittenhouse custom tages and that are not shared with outside providers-->
<!-- this file is included from ritthier2.redecl if identified in rittredeclmeta.mod -->

<!-- Definition of custom tag -->
<!ELEMENT risindex %ho; 	(risterm+,ristopic+,ristype+,risrule+,risposid+)>

<!ELEMENT risposid %ho; (#PCDATA)> 
<!ELEMENT risterm %ho; (#PCDATA)>
<!ELEMENT ristopic %ho; (#PCDATA)>
<!ELEMENT ristype %ho; (#PCDATA)>
<!ELEMENT risrule %ho; (#PCDATA)>
<!ELEMENT chapterid %ho; (#PCDATA)>
<!ELEMENT chaptertitle %ho; (#PCDATA)>

<!ELEMENT risinfo %ho; ((author| authorgroup| booktitle| chapternumber|  editor| isbn| mediaobject| primaryauthor| pubdate| publisher| riscurrent| risnext| risprev |  chaptertitle|chapterid)*)>

<!ELEMENT booktitle %ho; (#PCDATA)>
<!ELEMENT chapternumber %ho; (#PCDATA)>

<!ELEMENT primaryauthor %ho; ((personname|(%person.ident.mix;)+),(personblurb|email|address)*)>


<!ELEMENT riscurrent %ho; (#PCDATA)>
<!ELEMENT risnext %ho; (#PCDATA)>
<!ELEMENT risprev %ho; (#PCDATA)>



<!-- insertion in to sect1info -->
<!ENTITY % local.sect1info.attrib "">
<!ENTITY % sect1info.role.attrib "%role.attrib;">

<!ENTITY % sect1info.element "IGNORE">
<!ELEMENT sect1info %ho; (
(risindex|risinfo|%info.class;)+)
		%beginpage.exclusion;>
<!--end of sect1info.element-->

