<!-- rittmetaredeclare.mod -->
<!-- This file contains a list of modules that are used in redeclaration and there local values-->
<!-- the syntax currently assumes that all files are held loclly -->
<!-- this file holds the rittenhouse definitions for the following insertion points -->
<!-- pool.redecl, hier.redecl and hier2.redecl all of which are called ritt*.redecl -->

<!ENTITY % rittexclusions.module "INCLUDE">
  
<!-- include redefinitions file-->
<!ENTITY % rittredef.module "INCLUDE">

<!-- define first intermod for items between hier and pool -->
<!ENTITY % intermod.redecl.module "INCLUDE">
<![%intermod.redecl.module;[
<!ENTITY % rdbmods SYSTEM "rittintermod.mod">
]]>

<!-- define first redelcare for mixes -->
<!ENTITY % dbpool.redecl.module "INCLUDE">
<!ENTITY % rdbpool SYSTEM "rittredecl.mod">

<!-- define second hierarcy insertion point-->
<!ENTITY % dbhier.redecl.module "INCLUDE">
<![%dbhier.redecl.module;[
<!ENTITY % rdbhier SYSTEM "ritthier.mod">
]]>
 
<!-- define second hierarcy insertion point-->
<!ENTITY % dbhier.redecl2.module "INCLUDE">
<![%dbhier.redecl2.module;[
<!ENTITY % rdbhier2 SYSTEM "ritthier2.mod">
]]>

<!-- define custom insertion into second hierarcy insertion point-->
<!ENTITY %  rittcustomtags.redecl.module "INCLUDE">
<![%rittcustomtags.redecl.module;[
<!ENTITY % rittcustomtags SYSTEM "rittcustomtags.mod">
]]>

