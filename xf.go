package xls

type xf5 struct {
	Font      uint16
	Format    uint16
	Type      uint16
	Align     uint16
	Color     uint16
	Fill      uint16
	Border    uint16
	Linestyle uint16
}

func (x *xf5) formatNo() uint16 {
	return x.Format
}

type xf8 struct {
	Font        uint16
	Format      uint16
	Type        uint16
	Align       byte
	Rotation    byte
	Ident       byte
	Usedattr    byte
	Linestyle   uint32
	Linecolor   uint32
	Groundcolor uint16
}

func (x *xf8) formatNo() uint16 {
	return x.Format
}

type XF interface {
	formatNo() uint16
}
