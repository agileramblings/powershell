# <FIELD name="Stack Rank" refname="Microsoft.VSTS.Common.StackRank" type="Double" />
$newStackRank = $orgReg.CreateElement("FIELD")
$newStackRank.SetAttribute('name','Stack Rank')
$newStackRank.SetAttribute('refname','Microsoft.VSTS.Common.StackRank')
$newStackRank.SetAttribute('type','Double')
$orgReg.WITD.WORKITEMTYPE.FIELDS.AppendChild($newStackRank)

# <FIELD name="Size" refname="Microsoft.VSTS.Scheduling.Size" type="Integer" />
$newSize = $orgReg.CreateElement("FIELD")
$newSize.SetAttribute('name','Size')
$newSize.SetAttribute('refname','Microsoft.VSTS.Scheduling.Size')
$newSize.SetAttribute('type','Integer')
$orgReg.WITD.WORKITEMTYPE.FIELDS.AppendChild($newSize)
