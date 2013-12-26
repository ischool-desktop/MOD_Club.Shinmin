using System;
using FISCA.Permission;

namespace K12.Club.Shinmin
{
    public partial class DetailContentBase : FISCA.Presentation.DetailContent 
    {
        public DetailContentBase()
        {
            InitializeComponent();
        }

        public new bool SaveButtonVisible
        {
            get { return base.SaveButtonVisible; }
            set
            {
                //判斷權限
                if (Attribute.IsDefined(GetType(), typeof(FeatureCodeAttribute)))
                {
                    FeatureCodeAttribute fca = Attribute.GetCustomAttribute(GetType(), typeof(FeatureCodeAttribute)) as FeatureCodeAttribute;
                    if (fca != null)
                    {
                        if (FISCA.Permission.UserAcl.Current[fca.Code].Editable)
                            base.SaveButtonVisible = value;
                    }
                }
                else //沒有定議權限就按照平常方法處理。
                    base.SaveButtonVisible = value;
            }
        }

        public new bool CancelButtonVisible
        {
            get { return base.CancelButtonVisible; }
            set
            {
                //判斷權限
                if (Attribute.IsDefined(GetType(), typeof(FeatureCodeAttribute)))
                {
                    FeatureCodeAttribute fca = Attribute.GetCustomAttribute(GetType(), typeof(FeatureCodeAttribute)) as FeatureCodeAttribute;
                    if (fca != null)
                    {
                        if (FISCA.Permission.UserAcl.Current[fca.Code].Editable)
                            base.CancelButtonVisible = value;
                    }
                }
                else //沒有定議權限就按照平常方法處理。
                    base.CancelButtonVisible = value;
            }
        }

    }
}