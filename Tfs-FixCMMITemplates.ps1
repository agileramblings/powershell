# <FIELD name="Stack Rank" refname="Microsoft.VSTS.Common.StackRank" type="Double" />
$newStackRank = $orgReq.CreateElement("FIELD")
$newStackRank.SetAttribute('name','Stack Rank')
$newStackRank.SetAttribute('refname','Microsoft.VSTS.Common.StackRank')
$newStackRank.SetAttribute('type','Double')
$orgReq.WITD.WORKITEMTYPE.FIELDS.AppendChild($newStackRank)

# <FIELD name="Size" refname="Microsoft.VSTS.Scheduling.Size" type="Integer" />
$newSize = $orgReq.CreateElement("FIELD")
$newSize.SetAttribute('name','Size')
$newSize.SetAttribute('refname','Microsoft.VSTS.Scheduling.Size')
$newSize.SetAttribute('type','Integer')
$orgReq.WITD.WORKITEMTYPE.FIELDS.AppendChild($newSize)
