select
   fcomponent as part#,
   fcomprev as rev,
   fitem as item,
   fparent as parent,
   fparentrev as parent_rev,
   fqty as qty,
   fltooling as is_tooling,
   identity_column as id,
   fbommemo as memo,
   fnoperno as operation#
from inboms
 where
    fparent = '0042-65918' and fparentrev = '02'




