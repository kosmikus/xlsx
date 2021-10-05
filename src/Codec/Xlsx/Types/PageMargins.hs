{-# LANGUAGE CPP #-}
{-# LANGUAGE OverloadedStrings #-}
{-# LANGUAGE TemplateHaskell   #-}
{-# LANGUAGE RecordWildCards   #-}
{-# LANGUAGE DeriveGeneric #-}
module Codec.Xlsx.Types.PageMargins where

#ifdef USE_MICROLENS
import Lens.Micro.TH (makeLenses)
#else
import Control.Lens (makeLenses)
#endif

import Control.DeepSeq (NFData)
import qualified Data.Map as Map
import GHC.Generics (Generic)
import Text.XML

import Codec.Xlsx.Parser.Internal
import Codec.Xlsx.Writer.Internal

-- | See 18.3.1.62 "pageMargins"
data PageMargins = PageMargins {
    _pageMarginsLeft   :: Double
  , _pageMarginsRight  :: Double
  , _pageMarginsTop    :: Double
  , _pageMarginsBottom :: Double
  , _pageMarginsHeader :: Double
  , _pageMarginsFooter :: Double
  }
  deriving (Eq, Ord, Show, Generic)
instance NFData PageMargins

makeLenses ''PageMargins

instance ToElement PageMargins where
  toElement nm PageMargins{..} = Element {
      elementName  = nm
    , elementNodes = []
    , elementAttributes = Map.fromList $ [
          "left"   .= _pageMarginsLeft
        , "right"  .= _pageMarginsRight
        , "top"    .= _pageMarginsTop
        , "bottom" .= _pageMarginsBottom
        , "header" .= _pageMarginsHeader
        , "footer" .= _pageMarginsFooter
        ]
    }

instance FromCursor PageMargins where
  fromCursor cur = do
    _pageMarginsLeft   <- fromAttribute "left" cur
    _pageMarginsRight  <- fromAttribute "right" cur
    _pageMarginsTop    <- fromAttribute "top" cur
    _pageMarginsBottom <- fromAttribute "bottom" cur
    _pageMarginsHeader <- fromAttribute "header" cur
    _pageMarginsFooter <- fromAttribute "footer" cur
    return PageMargins{..}

instance FromXenoNode PageMargins where
  fromXenoNode root =
    parseAttributes root $ do
      _pageMarginsLeft   <- fromAttr "left"
      _pageMarginsRight  <- fromAttr "right"
      _pageMarginsTop    <- fromAttr "top"
      _pageMarginsBottom <- fromAttr "bottom"
      _pageMarginsHeader <- fromAttr "header"
      _pageMarginsFooter <- fromAttr "footer"
      return PageMargins{..}
