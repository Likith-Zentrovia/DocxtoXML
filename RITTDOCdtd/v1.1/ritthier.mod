<!-- ritthier.mod -->
<!-- This file contains a elements that are declared at the second hierarchy insertion point-->
<!-- element set in the Rittenhose variation of the DocBook DTD -->
<!-- redeclarations are preformed in many ways -->

<!-- check for inclusion of custom tags -->
<!-- remove the following items from parts ... |%refentry.class;|reference  -->
<!ENTITY % local.partcontent.mix "|subpart">
<!ENTITY % partcontent.mix
		"%appendix.class;|%chapter.class;|%nav.class;|%article.class;
		|preface %local.partcontent.mix;">

<!-- add subpart hierarcical entry between parts and chapters -->
<!ENTITY % subpart.module "INCLUDE">
<![%subpart.module;[

<!-- Note that subpart was to have its content model reduced in V4.3.  This
change will not be made after all. -->

<!ENTITY % local.subpart.attrib "">
<!ENTITY % subpart.role.attrib "%role.attrib;">

<!ENTITY % partcomponent.title.content
	"title, subtitle?, titleabbrev?">


<!ENTITY % subpart.element "INCLUDE">
<![%subpart.element;[
<!ELEMENT subpart %ho; (beginpage?,
                partinfo?, (%partcomponent.title.content;), partintro?,
		(%partcontent.mix;)+)
		%ubiq.inclusion;>
<!--end of subpart.element-->]]>

<!ENTITY % subpart.attlist "INCLUDE">
<![%subpart.attlist;[
<!ATTLIST subpart
		%label.attrib;
		%status.attrib;
		%common.attrib;
		%subpart.role.attrib;
		%local.subpart.attrib;
>
<!--end of subpart.attlist-->]]>
<!--end of subpart.module-->]]>

<!-- Dedication, ToC, and LoT ............................................. -->

<!ENTITY % dedication.module "IGNORE">


<!ENTITY % role.attrib
	"role		CDATA		#IMPLIED">

<!ENTITY % local.dedication.attrib "">
<!ENTITY % dedication.role.attrib "%role.attrib;">

<!ELEMENT dedication %ho; ( risinfo?, (%sect.title.content;)?, (%legalnotice.mix;)+)>
<!ATTLIST dedication
		%status.attrib;
		%common.attrib;
		%dedication.role.attrib;
		%local.dedication.attrib;
>
