﻿//------------------------------------------------------------------------------
// <auto-generated>
//     Этот код создан по шаблону.
//
//     Изменения, вносимые в этот файл вручную, могут привести к непредвиденной работе приложения.
//     Изменения, вносимые в этот файл вручную, будут перезаписаны при повторном создании кода.
// </auto-generated>
//------------------------------------------------------------------------------

namespace Tyreshop.DbAccess
{
    using System;
    using System.Data.Entity;
    using System.Data.Entity.Infrastructure;
    
    public partial class u0324292_mainEntities : DbContext
    {
        public u0324292_mainEntities()
            : base("name=u0324292_mainEntities")
        {
        }
    
        protected override void OnModelCreating(DbModelBuilder modelBuilder)
        {
            throw new UnintentionalCodeFirstException();
        }
    
        public virtual DbSet<car> cars { get; set; }
        public virtual DbSet<shop_address> shop_address { get; set; }
        public virtual DbSet<shop_affiliate> shop_affiliate { get; set; }
        public virtual DbSet<shop_affiliate_activity> shop_affiliate_activity { get; set; }
        public virtual DbSet<shop_affiliate_login> shop_affiliate_login { get; set; }
        public virtual DbSet<shop_affiliate_transaction> shop_affiliate_transaction { get; set; }
        public virtual DbSet<shop_api> shop_api { get; set; }
        public virtual DbSet<shop_api_ip> shop_api_ip { get; set; }
        public virtual DbSet<shop_api_session> shop_api_session { get; set; }
        public virtual DbSet<shop_attribute> shop_attribute { get; set; }
        public virtual DbSet<shop_attribute_description> shop_attribute_description { get; set; }
        public virtual DbSet<shop_attribute_group> shop_attribute_group { get; set; }
        public virtual DbSet<shop_attribute_group_description> shop_attribute_group_description { get; set; }
        public virtual DbSet<shop_banner> shop_banner { get; set; }
        public virtual DbSet<shop_banner_image> shop_banner_image { get; set; }
        public virtual DbSet<shop_banner_image_description> shop_banner_image_description { get; set; }
        public virtual DbSet<shop_cart> shop_cart { get; set; }
        public virtual DbSet<shop_category> shop_category { get; set; }
        public virtual DbSet<shop_category_description> shop_category_description { get; set; }
        public virtual DbSet<shop_category_filter> shop_category_filter { get; set; }
        public virtual DbSet<shop_category_path> shop_category_path { get; set; }
        public virtual DbSet<shop_category_to_layout> shop_category_to_layout { get; set; }
        public virtual DbSet<shop_category_to_store> shop_category_to_store { get; set; }
        public virtual DbSet<shop_comdb_summer> shop_comdb_summer { get; set; }
        public virtual DbSet<shop_comdb_winter> shop_comdb_winter { get; set; }
        public virtual DbSet<shop_country> shop_country { get; set; }
        public virtual DbSet<shop_coupon> shop_coupon { get; set; }
        public virtual DbSet<shop_coupon_category> shop_coupon_category { get; set; }
        public virtual DbSet<shop_coupon_history> shop_coupon_history { get; set; }
        public virtual DbSet<shop_coupon_product> shop_coupon_product { get; set; }
        public virtual DbSet<shop_currency> shop_currency { get; set; }
        public virtual DbSet<shop_custom_field> shop_custom_field { get; set; }
        public virtual DbSet<shop_custom_field_customer_group> shop_custom_field_customer_group { get; set; }
        public virtual DbSet<shop_custom_field_description> shop_custom_field_description { get; set; }
        public virtual DbSet<shop_custom_field_value> shop_custom_field_value { get; set; }
        public virtual DbSet<shop_custom_field_value_description> shop_custom_field_value_description { get; set; }
        public virtual DbSet<shop_customer> shop_customer { get; set; }
        public virtual DbSet<shop_customer_activity> shop_customer_activity { get; set; }
        public virtual DbSet<shop_customer_group> shop_customer_group { get; set; }
        public virtual DbSet<shop_customer_group_description> shop_customer_group_description { get; set; }
        public virtual DbSet<shop_customer_history> shop_customer_history { get; set; }
        public virtual DbSet<shop_customer_ip> shop_customer_ip { get; set; }
        public virtual DbSet<shop_customer_login> shop_customer_login { get; set; }
        public virtual DbSet<shop_customer_online> shop_customer_online { get; set; }
        public virtual DbSet<shop_customer_reward> shop_customer_reward { get; set; }
        public virtual DbSet<shop_customer_transaction> shop_customer_transaction { get; set; }
        public virtual DbSet<shop_customer_wishlist> shop_customer_wishlist { get; set; }
        public virtual DbSet<shop_disks> shop_disks { get; set; }
        public virtual DbSet<shop_download> shop_download { get; set; }
        public virtual DbSet<shop_download_description> shop_download_description { get; set; }
        public virtual DbSet<shop_dqc_setting> shop_dqc_setting { get; set; }
        public virtual DbSet<shop_event> shop_event { get; set; }
        public virtual DbSet<shop_extension> shop_extension { get; set; }
        public virtual DbSet<shop_filter> shop_filter { get; set; }
        public virtual DbSet<shop_filter_description> shop_filter_description { get; set; }
        public virtual DbSet<shop_filter_group> shop_filter_group { get; set; }
        public virtual DbSet<shop_filter_group_description> shop_filter_group_description { get; set; }
        public virtual DbSet<shop_geo_zone> shop_geo_zone { get; set; }
        public virtual DbSet<shop_gift_teaser> shop_gift_teaser { get; set; }
        public virtual DbSet<shop_google_base_category> shop_google_base_category { get; set; }
        public virtual DbSet<shop_google_base_category_to_category> shop_google_base_category_to_category { get; set; }
        public virtual DbSet<shop_information> shop_information { get; set; }
        public virtual DbSet<shop_information_description> shop_information_description { get; set; }
        public virtual DbSet<shop_information_to_layout> shop_information_to_layout { get; set; }
        public virtual DbSet<shop_information_to_store> shop_information_to_store { get; set; }
        public virtual DbSet<shop_language> shop_language { get; set; }
        public virtual DbSet<shop_layout> shop_layout { get; set; }
        public virtual DbSet<shop_layout_module> shop_layout_module { get; set; }
        public virtual DbSet<shop_layout_route> shop_layout_route { get; set; }
        public virtual DbSet<shop_length_class> shop_length_class { get; set; }
        public virtual DbSet<shop_length_class_description> shop_length_class_description { get; set; }
        public virtual DbSet<shop_location> shop_location { get; set; }
        public virtual DbSet<shop_manufacturer> shop_manufacturer { get; set; }
        public virtual DbSet<shop_manufacturer_to_store> shop_manufacturer_to_store { get; set; }
        public virtual DbSet<shop_marketing> shop_marketing { get; set; }
        public virtual DbSet<shop_modification> shop_modification { get; set; }
        public virtual DbSet<shop_module> shop_module { get; set; }
        public virtual DbSet<shop_mws_return> shop_mws_return { get; set; }
        public virtual DbSet<shop_option> shop_option { get; set; }
        public virtual DbSet<shop_option_description> shop_option_description { get; set; }
        public virtual DbSet<shop_option_value> shop_option_value { get; set; }
        public virtual DbSet<shop_option_value_description> shop_option_value_description { get; set; }
        public virtual DbSet<shop_order> shop_order { get; set; }
        public virtual DbSet<shop_order_custom_field> shop_order_custom_field { get; set; }
        public virtual DbSet<shop_order_history> shop_order_history { get; set; }
        public virtual DbSet<shop_order_option> shop_order_option { get; set; }
        public virtual DbSet<shop_order_product> shop_order_product { get; set; }
        public virtual DbSet<shop_order_recurring> shop_order_recurring { get; set; }
        public virtual DbSet<shop_order_recurring_transaction> shop_order_recurring_transaction { get; set; }
        public virtual DbSet<shop_order_status> shop_order_status { get; set; }
        public virtual DbSet<shop_order_total> shop_order_total { get; set; }
        public virtual DbSet<shop_order_voucher> shop_order_voucher { get; set; }
        public virtual DbSet<shop_pokupki_orders> shop_pokupki_orders { get; set; }
        public virtual DbSet<shop_product> shop_product { get; set; }
        public virtual DbSet<shop_product_attribute> shop_product_attribute { get; set; }
        public virtual DbSet<shop_product_description> shop_product_description { get; set; }
        public virtual DbSet<shop_product_discount> shop_product_discount { get; set; }
        public virtual DbSet<shop_product_filter> shop_product_filter { get; set; }
        public virtual DbSet<shop_product_image> shop_product_image { get; set; }
        public virtual DbSet<shop_product_option> shop_product_option { get; set; }
        public virtual DbSet<shop_product_option_value> shop_product_option_value { get; set; }
        public virtual DbSet<shop_product_recurring> shop_product_recurring { get; set; }
        public virtual DbSet<shop_product_related> shop_product_related { get; set; }
        public virtual DbSet<shop_product_reward> shop_product_reward { get; set; }
        public virtual DbSet<shop_product_special> shop_product_special { get; set; }
        public virtual DbSet<shop_product_to_category> shop_product_to_category { get; set; }
        public virtual DbSet<shop_product_to_download> shop_product_to_download { get; set; }
        public virtual DbSet<shop_product_to_layout> shop_product_to_layout { get; set; }
        public virtual DbSet<shop_product_to_store> shop_product_to_store { get; set; }
        public virtual DbSet<shop_recurring> shop_recurring { get; set; }
        public virtual DbSet<shop_recurring_description> shop_recurring_description { get; set; }
        public virtual DbSet<shop_return> shop_return { get; set; }
        public virtual DbSet<shop_return_action> shop_return_action { get; set; }
        public virtual DbSet<shop_return_history> shop_return_history { get; set; }
        public virtual DbSet<shop_return_reason> shop_return_reason { get; set; }
        public virtual DbSet<shop_return_status> shop_return_status { get; set; }
        public virtual DbSet<shop_review> shop_review { get; set; }
        public virtual DbSet<shop_setting> shop_setting { get; set; }
        public virtual DbSet<shop_stock_status> shop_stock_status { get; set; }
        public virtual DbSet<shop_store> shop_store { get; set; }
        public virtual DbSet<shop_tax_class> shop_tax_class { get; set; }
        public virtual DbSet<shop_tax_rate> shop_tax_rate { get; set; }
        public virtual DbSet<shop_tax_rate_to_customer_group> shop_tax_rate_to_customer_group { get; set; }
        public virtual DbSet<shop_tax_rule> shop_tax_rule { get; set; }
        public virtual DbSet<shop_tdb_summer> shop_tdb_summer { get; set; }
        public virtual DbSet<shop_tdb_winter> shop_tdb_winter { get; set; }
        public virtual DbSet<shop_upload> shop_upload { get; set; }
        public virtual DbSet<shop_url_alias> shop_url_alias { get; set; }
        public virtual DbSet<shop_user> shop_user { get; set; }
        public virtual DbSet<shop_user_group> shop_user_group { get; set; }
        public virtual DbSet<shop_voucher> shop_voucher { get; set; }
        public virtual DbSet<shop_voucher_history> shop_voucher_history { get; set; }
        public virtual DbSet<shop_voucher_theme> shop_voucher_theme { get; set; }
        public virtual DbSet<shop_voucher_theme_description> shop_voucher_theme_description { get; set; }
        public virtual DbSet<shop_weight_class> shop_weight_class { get; set; }
        public virtual DbSet<shop_weight_class_description> shop_weight_class_description { get; set; }
        public virtual DbSet<shop_yml> shop_yml { get; set; }
        public virtual DbSet<shop_zone> shop_zone { get; set; }
        public virtual DbSet<shop_zone_to_geo_zone> shop_zone_to_geo_zone { get; set; }
        public virtual DbSet<tsrim> tsrims { get; set; }
    }
}
