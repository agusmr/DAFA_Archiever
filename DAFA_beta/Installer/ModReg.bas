$location == 'right')
            $this->_text_h_align = 3; 
        if ($location == 'fill')
            $this->_text_h_align = 4; 
        if ($location == 'justify')
            $this->_text_h_align = 5;
        if ($location == 'merge')
            $this->_text_h_align = 6;
        if ($location == 'equal_space') // For T.K.
            $this->_text_h_align = 7; 
        if ($location == 'top')
            $this->text_v_align = 0; 
        if ($location == 'vcentre')
            $this->text_v_align = 1; 
        if ($location == 'vcenter')
            $this->text_v_align = 1; 
        if ($location == 'bottom')
            $this->text_v_align = 2; 
        if ($location == 'vjustify')
            $this->text_v_align = 3; 
        if ($location == 'vequal_space') // For T.K.
            $this->text_v_align = 4; 
    }
    
    /**
    * This is an alias for the unintuitive set_align('merge')
    *
    * @access public
    */
    function set_merge()
    {
        $this->set_align('merge');
    }
    
    /**
    * Bold has a range 0x64..0x3E8.
    * 0x190 is normal. 0x2BC is bold.
    *
    * @access public
    * @param integer $weight Weight for the